Attribute VB_Name = "GlHoraires"
Option Explicit

Public Type Couleur_Type
    Fond As Long
    Titre As Long
    HeureCourante As Long
    ProchaineVisite As Long
    AutreVisite As Long
    EtatVisite As Long
End Type

Public Enum Etat_Enum
    Ouvert = 1
    Complet = 2
    Ferme = 3
End Enum

Public Type Option_Type
    NbVisites   As Long             ' Nb de visites affich�es simultan�ment
    DureeVisite As Long             ' Dur�e d'une visite (en minutes)
    IntervalleMAJHeure As Long      ' Intervalle entre 2 MAJ de l'heure courante (en millisecondes)
    IntervalleClignote As Long      ' P�riode de clignotement de l'etat (en millisecondes)
    Couleurs As Couleur_Type        ' Couleurs utilis�es
End Type

Public gOptions As Option_Type

Public gTabHeureVisite() As Date
Public gTabNoVisite() As Long
Public gTabEtatVisite() As Etat_Enum
