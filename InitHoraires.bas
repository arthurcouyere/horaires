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
    Dim cFichier As String

    With frmHoraires
        
        ' Timer
        .tmrHoraires.Interval = gOptions.IntervalleMAJHeure
        .tmrHoraires.Enabled = True

        ' Construction dynamique de la liste des visites
        For i = .lblHeureVisite.LBound + 1 To gOptions.NbVisites
            
            ' 1ere visite suivante
            If i = .lblHeureVisite.LBound + 1 Then
                Load .lblVisite(i)
                .lblVisite(i).Visible = True
                .lblVisite(i).Top = .lblVisite(i - 1).Top + .lblVisite(i - 1).Height + ESPACE_VISITES
                .lblVisite(i).Caption = LIBELLE_AUTRE_VISITE
                .lblVisite(i).FontBold = False
                .lblVisite(i).FontSize = 40
                .lblVisite(i).Height = 1095
                
            End If
                
            Load .lblHeureVisite(i)
            .lblHeureVisite(i).Visible = True
            .lblHeureVisite(i).Top = .lblHeureVisite(i - 1).Top + .lblHeureVisite(i - 1).Height + ESPACE_VISITES
            .lblHeureVisite(i).FontBold = False
            .lblHeureVisite(i).FontSize = 40
            .lblHeureVisite(i).Height = 1095
            
            Load .lblNoVisite(i)
            .lblNoVisite(i).Visible = True
            .lblNoVisite(i).Top = .lblHeureVisite(i).Top
            .lblNoVisite(i).FontBold = False
            .lblNoVisite(i).FontSize = 40
            .lblNoVisite(i).Height = 1095
            
        Next
        
        ' Couleurs
        .BackColor = gOptions.Couleurs.Fond
        
        .picHoraires.BackColor = gOptions.Couleurs.Fond
        .lblTitre.BackColor = gOptions.Couleurs.Fond
        .lblTitre.ForeColor = gOptions.Couleurs.Titre
        
        .lblHeureCourante.BackColor = gOptions.Couleurs.Fond
        .lblHeureCourante.ForeColor = gOptions.Couleurs.HeureCourante
        .lblHeure.BackColor = gOptions.Couleurs.Fond
        .lblHeure.ForeColor = gOptions.Couleurs.HeureCourante
        
        For i = .lblHeureVisite.LBound To .lblHeureVisite.UBound
            If i <= 2 Then .lblVisite(i).BackColor = gOptions.Couleurs.Fond
            .lblHeureVisite(i).BackColor = gOptions.Couleurs.Fond
            .lblNoVisite(i).BackColor = gOptions.Couleurs.Fond
            If i = .lblHeureVisite.LBound Then
                .lblVisite(i).ForeColor = gOptions.Couleurs.ProchaineVisite
                .lblHeureVisite(i).ForeColor = gOptions.Couleurs.ProchaineVisite
                .lblNoVisite(i).ForeColor = gOptions.Couleurs.ProchaineVisite
            Else
                If i <= 2 Then .lblVisite(i).ForeColor = gOptions.Couleurs.AutreVisite
                .lblHeureVisite(i).ForeColor = gOptions.Couleurs.AutreVisite
                .lblNoVisite(i).ForeColor = gOptions.Couleurs.AutreVisite
            End If
        Next
        
        ' Image
        cFichier = App.Path & "\" & FICHIER_IMAGE
        Set .picHoraires.Picture = LoadPicture(cFichier)
        Set .Palette = LoadPicture(cFichier)
        .PaletteMode = vbPaletteModeCustom
        
    End With
    
End Sub


'----------------------------------------------------------------
' But : Initialise le tableau des horaires des visites
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
    For i = 1 To gOptions.NbVisites
        gTabHeureVisite(i) = DateAdd("n", gOptions.DureeVisite * i, dTime)
    Next
    
End Sub


'----------------------------------------------------------------
' But : Initialise le tableau des numéros de visite
' Entrées :
' Sorties : tableau gTabNoVisite mis à jour
' Suppositions :
'       Les paramètres généraux ont été chargés
' Effets de bord : IHM
'----------------------------------------------------------------
Public Sub InitTabNoVisite()
    Dim i As Long
    
    ReDim gTabNoVisite(1 To gOptions.NbVisites)
    For i = 1 To gOptions.NbVisites
        gTabNoVisite(i) = i
    Next
    
End Sub




