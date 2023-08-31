VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDuree 
      Height          =   285
      Left            =   2715
      MaxLength       =   3
      TabIndex        =   0
      Top             =   465
      Width           =   495
   End
   Begin VB.CommandButton cmdAnnuler 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   2235
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   435
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblMinutes 
      Caption         =   "min"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblDureeVisite 
      Caption         =   "Intervalle entre deux visites"
      Height          =   255
      Left            =   555
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnnuler_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo TraitementErreur
    
    gOptions.DureeVisite = CLng(txtDuree.Text)
    Call SaveSetting(App.Title, "Général", "DuréeVisite", gOptions.DureeVisite)

TraitementErreur:

    Unload Me

End Sub

Private Sub Form_Load()
    txtDuree.Text = gOptions.DureeVisite
End Sub


Private Sub txtDuree_GotFocus()

    txtDuree.SelStart = 0
    txtDuree.SelLength = Len(txtDuree.Text)
    
End Sub
