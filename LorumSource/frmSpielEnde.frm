VERSION 5.00
Begin VB.Form frmSpielEnde 
   Caption         =   "Spiel-Ende"
   ClientHeight    =   3012
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3012
   ScaleWidth      =   3744
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnOK 
      Caption         =   "Spiel fortsetzen"
      Default         =   -1  'True
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   3492
   End
   Begin VB.CommandButton btnSpielWiederholen 
      Caption         =   "Spiel wiederholen"
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3492
   End
   Begin VB.Label lblSpielerPunkte 
      Caption         =   "99"
      Height          =   372
      Left            =   960
      TabIndex        =   8
      Top             =   2040
      Width           =   372
   End
   Begin VB.Label lblRechtsPunkte 
      Caption         =   "99"
      Height          =   372
      Left            =   960
      TabIndex        =   7
      Top             =   1560
      Width           =   372
   End
   Begin VB.Label lblMittepunkte 
      Caption         =   "99"
      Height          =   372
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   372
   End
   Begin VB.Label lblLinksPunkte 
      Caption         =   "99"
      Height          =   372
      Left            =   960
      TabIndex        =   5
      Top             =   600
      Width           =   372
   End
   Begin VB.Label lblSpieler 
      Caption         =   "Punkte Spieler"
      Height          =   372
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   1932
   End
   Begin VB.Label lblRechts 
      Caption         =   "Punkte Rechts"
      Height          =   372
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   1932
   End
   Begin VB.Label lblMitte 
      Caption         =   "Punkte Mitte"
      Height          =   372
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1932
   End
   Begin VB.Label lblLinks 
      Caption         =   "Punkte Links"
      Height          =   372
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1932
   End
End
Attribute VB_Name = "frmSpielEnde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    blnSpielWiederholen = False
    Unload Me
End Sub

Private Sub btnSpielWiederholen_Click()
    blnSpielWiederholen = True
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)
    lblLinksPunkte.Caption = frmRoserl.LinksPunkte
    lblMittepunkte = frmRoserl.MittePunkte
    lblRechtsPunkte = frmRoserl.RechtsPunkte
    lblSpielerPunkte = frmRoserl.SpielerPunkte
    #If Spielwiederholen = 0 Then                   '0=nein                         'Gerbing 19.05.2021
        btnSpielWiederholen.Visible = False
    #Else
        btnSpielWiederholen.Visible = True
    #End If
End Sub
