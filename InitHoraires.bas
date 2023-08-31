Attribute VB_Name = "InitHoraires"
Option Explicit


'----------------------------------------------------------------
' But : Initialise la fenêtre principale
' Entrées :
' Sorties :
' Suppositions :
'       Les paramètres généraux ont été chargés
' Effets de bord : IHM
'----------------------------------------------------------------
Public Sub InitFrmHoraires()
    Dim i As Long

    With frmHoraires
        
        ' Timer
        .tmrHoraires.Interval = gOptions.IntervalleMAJHeure
        .tmrHoraires.Enabled = True

        ' Construction dynamique de la liste des visites
        For i = .lblHeureVisite.LBound + 1 To gOptions.NbVisites
            If i = .lblHeureVisite.LBound + 1 Then
                Load .lblVisite(i)
                .lblVisite(i).Visible = True
                .lblVisite(i).Top = .lblVisite(i - 1).Top + .lblVisite(i - 1).Height + ESPACE_VISITES
                .lblVisite(i).Caption = LIBELLE_AUTRE_VISITE
                .lblVisite(i).FontBold = False
            End If
            
            Load .lblHeureVisite(i)
            .lblHeureVisite(i).Visible = True
            .lblHeureVisite(i).Top = .lblHeureVisite(i - 1).Top + .lblHeureVisite(i - 1).Height + ESPACE_VISITES
            .lblHeureVisite(i).FontBold = False
        Next
        
        ' Couleurs
        .BackColor = gOptions.Couleurs.Fond
        
        .picHoraires.BackColor = gOptions.Couleurs.Fond
        .lblTitre.BackColor = gOptions.Couleurs.Fond
        .lblTitre.ForeColor = gOptions.Couleurs.Titre
        
        .lblHeureCourante.BackColor = gOptions.Couleurs.Fond
        .lblHeureCourante.ForeColor = gOptions.Couleurs.HeureCourante
        
        For i = .lblHeureVisite.LBound To .lblHeureVisite.UBound
            If i <= 2 Then .lblVisite(i).BackColor = gOptions.Couleurs.Fond
            .lblHeureVisite(i).BackColor = gOptions.Couleurs.Fond
            If i = .lblHeureVisite.LBound Then
                .lblVisite(i).ForeColor = gOptions.Couleurs.ProchaineVisite
                .lblHeureVisite(i).ForeColor = gOptions.Couleurs.ProchaineVisite
            Else
                If i <= 2 Then .lblVisite(i).ForeColor = gOptions.Couleurs.AutreVisite
                .lblHeureVisite(i).ForeColor = gOptions.Couleurs.AutreVisite
            End If
        Next
        
        ' Image
        .picHoraires.Picture = LoadPicture(App.Path & "\Fond.jpg")
        
    End With
    
End Sub


'----------------------------------------------------------------
' But : Initialise le tableau des visites
' Entrées :
' Sorties : tableau gTabHeureVisite mis à jour
' Suppositions :
'       Les paramètres généraux ont été chargés
' Effets de bord : IHM
'----------------------------------------------------------------
Public Sub InitTabHeureVisite()
    Dim i As Long
    Dim dTime As Date
    
    dTime = CDate(Format(Time, "hh:nn"))
    
    ReDim gTabHeureVisite(1 To gOptions.NbVisites)
    For i = LBound(gTabHeureVisite) To UBound(gTabHeureVisite)
        gTabHeureVisite(i) = DateAdd("n", gOptions.DureeVisite * i, dTime)
    Next
    
End Sub
