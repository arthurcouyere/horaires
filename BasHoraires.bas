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
' Entr�es :
' Sorties :
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Function ArreterWindows() As Long

    Call ExitWindowsEx(EWX_SHUTDOWN + EWX_FORCE, 0)

End Function

'----------------------------------------------------------------
' But : Transforme une chaine de couleur RGB au format "RRGGBB"
'       en un long repr�sentant le code RGB
' Entr�es : Chaine
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
' But : R�cup�re les param�tres g�n�raux de l'application
' Entr�es :
' Sorties : variable gOptions mise � jour
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
        .IntervalleMAJHeure = GetSetting(App.Title, "G�n�ral", _
                                         "IntervalleMAJHeure", 50)
        .IntervalleClignote = GetSetting(App.Title, "G�n�ral", _
                                         "IntervalleClignote", 1000)
        
        ' Visites
        .DureeVisite = GetSetting(App.Title, "G�n�ral", "Dur�eVisite", 6)
        .NbVisites = NB_VISITES
        
    End With
    
End Sub

'----------------------------------------------------------------
' But : Decale une visite de x minutes
'       Si x est n�gatif, v�rifie que la nouvelle heure n'est pas
'       ant�rieure � l'heure courante
' Entr�es : Minutes, optionel Bool�en: Force (True) le d�calage
' Sorties :
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Sub DecaleVisite(Minute As Long, Optional Force As Boolean = False)
    Dim i As Long
    
    ' V�rifie que la nouvelle heure n'est pas ant�rieure � l'heure courante
    If DateAdd("n", Minute, gTabHeureVisite(1)) < Time And Not Force Then Exit Sub
    
    For i = 1 To gOptions.NbVisites
       gTabHeureVisite(i) = DateAdd("n", Minute, gTabHeureVisite(i))
    Next
    
End Sub

'----------------------------------------------------------------
' But : Decale les num�ros de visite de x
'       Si x est n�gatif, v�rifie que le nouveau num�ro n'est
'       pas inf�rieur � 1
' Entr�es : Nombre
' Sorties :
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Sub DecaleNoVisite(Nb As Long)
    Dim i As Long
    
    ' V�rifie que le nouveau num�ro n'est pas inf�rieur � 1
    If gTabNoVisite(1) + Nb < 1 Then Exit Sub
    
    For i = 1 To gOptions.NbVisites
       gTabNoVisite(i) = gTabNoVisite(i) + Nb
    Next
    
End Sub

'----------------------------------------------------------------
' But : Decale les etats de visite (une visite vient de commencer)
' Entr�es :
' Sorties :
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Sub DecaleEtatVisite()
    Dim i As Long
    
    For i = 1 To gOptions.NbVisites - 1
       gTabEtatVisite(i) = gTabEtatVisite(i + 1)
    Next
    
    ' Etat par d�faut de la derni�re visite (nouvellement ajout�e)
    gTabEtatVisite(gOptions.NbVisites) = Ouvert
    
End Sub

'----------------------------------------------------------------
' But : Modifie l'�tat de visite selon le cycle suivant:
'
'      +--> ouvert --> complet --> ferm� --+
'      |                                   |
'      +----<----------<----------<--------+
'
' Entr�es : No de la visite
' Sorties :
' Suppositions :
' Effets de bord : IHM
'----------------------------------------------------------------
Public Sub ModifEtatVisite(NumVisite As Long)

    ' V�rifie la validit� du num�ro de visite
    If NumVisite < 1 Or NumVisite > gOptions.NbVisites Then Exit Sub
    
    gTabEtatVisite(NumVisite) = (gTabEtatVisite(NumVisite)) Mod 3 + 1
End Sub

'----------------------------------------------------------------
' But : Met � jour le tableau des visites (cas de la modif de la
'       dur�e des visites)
' Entr�es :
' Sorties : tableau gTabHeureVisite mis � jour
' Suppositions :
'       Les param�tres g�n�raux ont �t� charg�s
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
' But : Renvoi le libell� d'un �tat de visite (ouvert / complet /ferm�)
' Entr�es : Etat
' Sorties : Libell�
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

