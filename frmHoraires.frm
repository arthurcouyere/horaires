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
   Begin VB.PictureBox picHoraires 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   600
      ScaleHeight     =   6015
      ScaleWidth      =   7695
      TabIndex        =   4
      Top             =   2640
      Width           =   7695
   End
   Begin VB.Timer tmrHoraires 
      Enabled         =   0   'False
      Left            =   240
      Top             =   120
   End
   Begin VB.Label lblTitre 
      Alignment       =   2  'Center
      Caption         =   "Horaires des visites"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   13575
   End
   Begin VB.Label lblHeureVisite 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   7920
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label lblVisite 
      Caption         =   "Prochaine visite"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label lblHeureCourante 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3960
      TabIndex        =   0
      Top             =   7800
      Width           =   7215
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
    
    Select Case KeyCode
    Case vbKeyEscape    ' Echap
        lRet = MsgBox("Voulez-vous quitter l'application ?", _
                       vbQuestion + vbOKCancel + vbDefaultButton2, "Horaires")
        If lRet = vbOK Then Unload Me
        
    Case vbKeyAdd       ' Plus
        Call DecaleVisite(1)
        Call PutVisitesDansForm
        
    Case vbKeySubtract  ' Moins
        Call DecaleVisite(-1)
        Call PutVisitesDansForm
    
    End Select
    
End Sub

Private Sub Form_Load()
    
    Call GetOptions
    Call InitFrmHoraires
    Call InitTabHeureVisite
    Call PutVisitesDansForm
    
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    ' Titre
    lblTitre.Left = (Me.Width - lblTitre.Width) / 2
    
    ' Heure courante
    lblHeureCourante.Left = (Me.Width - lblHeureCourante.Width) / 2
    lblHeureCourante.Top = Me.Height - lblHeureCourante.Height - 50
    
    ' Visites
    For i = lblHeureVisite.LBound To lblHeureVisite.UBound
        lblHeureVisite(i).Left = Me.Width - lblHeureVisite(i).Width - lblVisite(1).Left
    Next
    
    ' Images
    picHoraires.Left = lblVisite(2).Left
    picHoraires.Width = lblHeureVisite(2).Left - picHoraires.Left - 200
    picHoraires.Top = lblVisite(2).Top + lblVisite(2).Height + 50
    picHoraires.Height = lblHeureCourante.Top - picHoraires.Top
    
End Sub

Private Sub tmrHoraires_Timer()
    
    ' Affiche l'heure courante
    Call PutHeureCouranteDansLabel(Me.lblHeureCourante)
    
    ' Décale les visites si une visite vient de commencer
    If Time >= gTabHeureVisite(1) Then
        Call DecaleVisite(gOptions.DureeVisite)
        Call PutVisitesDansForm
    End If
    
End Sub
