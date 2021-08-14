VERSION 5.00
Object = "{54A43060-C747-4026-8E08-2A801095C838}#1.0#0"; "roserl.ocx"
Begin VB.Form frmLetztenStichZeigen 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Letzter Stich"
   ClientHeight    =   10476
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   10224
   Icon            =   "frmLetztenStichZeigen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   873
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   852
   StartUpPosition =   1  'Fenstermitte
   Begin RoserlOcx.Roserl SpielerGespielt 
      Height          =   3012
      Left            =   4320
      TabIndex        =   0
      Top             =   7200
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   5313
      FaceMode        =   0
      CardType        =   3
   End
   Begin RoserlOcx.Roserl RechtsGespielt 
      Height          =   3012
      Left            =   8280
      TabIndex        =   1
      Top             =   3720
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   5313
      FaceMode        =   0
      CardType        =   3
   End
   Begin RoserlOcx.Roserl MitteGespielt 
      Height          =   3012
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   5313
      FaceMode        =   0
      CardType        =   3
   End
   Begin RoserlOcx.Roserl LinksGespielt 
      Height          =   3012
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   5313
      FaceMode        =   0
      CardType        =   3
   End
End
Attribute VB_Name = "frmLetztenStichZeigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Rest As Integer
    Dim I As Integer
    
    Rest = AKIndex Mod 4                            'ich will den Rest sehen
    For I = AKIndex - 3 - Rest To AKIndex - Rest
        If AusgespielteKarten(I).cGespieltVon = Links Then
            LinksGespielt.CardType = AusgespielteKarten(I).cType            'Gerbing 29.07.2017
            LinksGespielt.CardValue = AusgespielteKarten(I).cValue
        ElseIf AusgespielteKarten(I).cGespieltVon = Mitte Then
            MitteGespielt.CardType = AusgespielteKarten(I).cType            'Gerbing 29.07.2017
            MitteGespielt.CardValue = AusgespielteKarten(I).cValue
        ElseIf AusgespielteKarten(I).cGespieltVon = Rechts Then
            RechtsGespielt.CardType = AusgespielteKarten(I).cType           'Gerbing 29.07.2017
            RechtsGespielt.CardValue = AusgespielteKarten(I).cValue
        ElseIf AusgespielteKarten(I).cGespieltVon = Spieler Then
            SpielerGespielt.CardType = AusgespielteKarten(I).cType          'Gerbing 29.07.2017
            SpielerGespielt.CardValue = AusgespielteKarten(I).cValue
        End If
    Next I
End Sub

