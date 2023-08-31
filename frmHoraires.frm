VERSION 5.00
Begin VB.Form frmHoraires 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14655
   Icon            =   "frmHoraires.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTitre 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   1080
      ScaleHeight     =   1455
      ScaleWidth      =   12255
      TabIndex        =   6
      Top             =   0
      Width           =   12255
   End
   Begin VB.Timer tmrHoraires 
      Enabled         =   0   'False
      Left            =   240
      Top             =   120
   End
   Begin VB.PictureBox picHoraires 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   600
      ScaleHeight     =   3855
      ScaleWidth      =   7575
      TabIndex        =   3
      Top             =   3360
      Width           =   7575
   End
   Begin VB.Label lblEtat 
      Caption         =   "Complet"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   1
      Left            =   5280
      TabIndex        =   7
      Top             =   1500
      Width           =   3015
   End
   Begin VB.Label lblNoVisite 
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   8280
      TabIndex        =   5
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label lblHeure 
      Caption         =   "Il est "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label lblHeureVisite 
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   10080
      TabIndex        =   2
      Top             =   1410
      Width           =   2520
   End
   Begin VB.Label lblVisite 
      Caption         =   "Prochaine visite"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Label lblHeureCourante 
      AutoSize        =   -1  'True
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   3960
      TabIndex        =   0
      Top             =   7920
      Width           =   5460
   End
End
Attribute VB_Name = "frmHoraires"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
' But:
'   Affichage des horaires de visite
'
' Auteur:
'   Emmanuel Jammes
'
' Version:
'   1.0
'
' Historique:
'
'   EJ  20/03/2001  1.0 Création
'
'----------------------------------------------------------------

Option Explicit


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRet As Long
    Dim lNumVisite As Long
    
    Select Case KeyCode
    Case vbKeyEscape    ' Echap
        lRet = MsgBox("Voulez-vous quitter l'application ?", _
                       vbQuestion + vbOKCancel, "Horaires")
        If lRet = vbOK Then Unload Me
        
    Case vbKeyDecimal    ' Point décimal
        lRet = MsgBox("Voulez-vous arrêter l'ordinateur ?", _
                       vbQuestion + vbOKCancel, "Horaires")
        If lRet = vbOK Then ArreterWindows
        
    Case vbKeyReturn    ' Entrée
        frmOptions.Show vbModal
        
        ' Met à jour les visites
        Call MajTabHeureVisite
        Call PutVisitesDansForm
        
    Case vbKeyAdd       ' Plus
        Call DecaleVisite(1)
        Call PutVisitesDansForm
        
    Case vbKeySubtract  ' Moins
        Call DecaleVisite(-1)
        Call PutVisitesDansForm
    
    Case vbKeyMultiply  ' Multiplier
        Call DecaleNoVisite(1)
        Call PutVisitesDansForm
    
    Case vbKeyDivide    ' Diviser
        Call DecaleNoVisite(-1)
        Call PutVisitesDansForm
    
    Case vbKeyNumpad0, _
         vbKeyNumpad1, _
         vbKeyNumpad2, _
         vbKeyNumpad3, _
         vbKeyNumpad4, _
         vbKeyNumpad5, _
         vbKeyNumpad6, _
         vbKeyNumpad7, _
         vbKeyNumpad8, _
         vbKeyNumpad9

         
        ' Numéro de visite :
        '    - 1 à 10 pour les touches 0 à 9 du pavé numérique
        lNumVisite = KeyCode - vbKeyNumpad0 + 1
        
        Call ModifEtatVisite(lNumVisite)
        Call PutVisitesDansForm
        
    End Select
    
End Sub

Private Sub Form_Load()
    
    Call GetOptions
    Call InitRegistry
    Call InitFrmHoraires
    Call InitTabHeureVisite
    Call InitTabNoVisite
    Call InitTabEtatVisite
    Call PutVisitesDansForm
    
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    ' Titre
    'lblTitre.Left = (Me.Width - lblTitre.Width) / 2
    picTitre.Left = (Me.Width - picTitre.Width) / 2
    
    ' Heure courante
    lblHeureCourante.Left = (Me.Width - lblHeureCourante.Width) / 2
    lblHeureCourante.Top = Me.Height - lblHeureCourante.Height - 50
    lblHeure.Left = lblHeureCourante.Left - lblHeure.Width
    lblHeure.Top = lblHeureCourante.Top + (lblHeureCourante.Height - lblHeure.Height) / 2
    
    ' Visites
    For i = lblHeureVisite.LBound To lblHeureVisite.UBound
        lblHeureVisite(i).Left = Me.Width - lblHeureVisite(i).Width - 200
        lblNoVisite(i).Left = lblHeureVisite(i).Left - lblNoVisite(i).Width - 200
        lblEtat(i).Left = lblNoVisite(i).Left - lblEtat(i).Width - 200
    Next
    
    ' Images
    picHoraires.Left = lblVisite(2).Left
    picHoraires.Width = lblNoVisite(2).Left - picHoraires.Left - 200
    picHoraires.Top = lblVisite(2).Top + lblVisite(2).Height + 50
    picHoraires.Height = lblHeureCourante.Top - picHoraires.Top
    
End Sub

Private Sub tmrHoraires_Timer()
    
    ' Affiche l'heure courante
    Call PutHeureCouranteDansLabel(Me.lblHeureCourante)
    
    ' Décale les visites si une visite vient de commencer
    If Time >= gTabHeureVisite(1) Then
        Call DecaleVisite(gOptions.DureeVisite)
        Call DecaleNoVisite(1)
        Call DecaleEtatVisite
        Call PutVisitesDansForm
    End If
    
End Sub
