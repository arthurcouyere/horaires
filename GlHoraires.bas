Attribute VB_Name = "GlHoraires"
Option Explicit

Public Type Couleur_Type
    Fond As Long
    Titre As Long
    HeureCourante As Long
    ProchaineVisite As Long
    AutreVisite As Long
End Type


Public Type Option_Type
    NbVisites   As Long             ' Nb de visites affich�es simultan�ment
    DureeVisite As Long             ' Dur�e d'une visite (en minutes)
    IntervalleMAJHeure As Long      ' Intervalle entre 2 MAJ de l'heure courante (en secondes)
    Couleurs As Couleur_Type        ' Couleurs utilis�es
End Type

Public gOptions As Option_Type

Public gTabHeureVisite() As Date
Public gTabNoVisite() As Long

