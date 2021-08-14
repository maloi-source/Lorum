VERSION 5.00
Begin VB.Form frmKarteVerschieben 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Karte verschieben"
   ClientHeight    =   3528
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4596
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3528
   ScaleWidth      =   4596
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   492
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   1932
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wohin soll die Karte verschoben werden"
      Height          =   2412
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4332
      Begin VB.OptionButton OptSpieler 
         Caption         =   "Spieler"
         Height          =   372
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   1692
      End
      Begin VB.OptionButton OptRechts 
         Caption         =   "Rechts"
         Height          =   372
         Left            =   600
         TabIndex        =   3
         Top             =   1320
         Width           =   1692
      End
      Begin VB.OptionButton OptMitte 
         Caption         =   "Mitte"
         Height          =   372
         Left            =   600
         TabIndex        =   2
         Top             =   840
         Width           =   1812
      End
      Begin VB.OptionButton OptLinks 
         Caption         =   "Links"
         Height          =   372
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   1692
      End
   End
End
Attribute VB_Name = "frmKarteVerschieben"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    If OptLinks.Value = True Then
        ReDim Preserve LinksCards(1 To UBound(LinksCards) + 1)
        Load frmRoserl.cdLinks(frmRoserl.cdLinks.Count)                             'Karte neu aufnehmen
        frmRoserl.cdLinks(frmRoserl.cdLinks.Count - 1).CardType = gintTypeVerschieben
        frmRoserl.cdLinks(frmRoserl.cdLinks.Count - 1).CardValue = gintValueVerschieben
        Call frmRoserl.LinksKartenNeu
        Call frmRoserl.HintersteEntladen(PlayerTurn, gintIndexVerschieben + 1)
    ElseIf OptMitte.Value = True Then
        ReDim Preserve MitteCards(1 To UBound(MitteCards) + 1)
        Load frmRoserl.cdMitte(frmRoserl.cdMitte.Count)
        frmRoserl.cdMitte(frmRoserl.cdMitte.Count - 1).CardType = gintTypeVerschieben
        frmRoserl.cdMitte(frmRoserl.cdMitte.Count - 1).CardValue = gintValueVerschieben
        Call frmRoserl.MitteKartenNeu
        Call frmRoserl.HintersteEntladen(PlayerTurn, gintIndexVerschieben + 1)
    ElseIf OptRechts.Value = True Then
        ReDim Preserve RechtsCards(1 To UBound(RechtsCards) + 1)
        Load frmRoserl.cdRechts(frmRoserl.cdRechts.Count)
        frmRoserl.cdRechts(frmRoserl.cdRechts.Count - 1).CardType = gintTypeVerschieben
        frmRoserl.cdRechts(frmRoserl.cdRechts.Count - 1).CardValue = gintValueVerschieben
        Call frmRoserl.RechtsKartenNeu
        Call frmRoserl.HintersteEntladen(PlayerTurn, gintIndexVerschieben + 1)
    ElseIf OptSpieler.Value = True Then
        ReDim Preserve SpielerCards(1 To UBound(SpielerCards) + 1)
        Load frmRoserl.cdSpieler(frmRoserl.cdSpieler.Count)
        frmRoserl.cdSpieler(frmRoserl.cdSpieler.Count - 1).CardType = gintTypeVerschieben
        frmRoserl.cdSpieler(frmRoserl.cdSpieler.Count - 1).CardValue = gintValueVerschieben
        Call frmRoserl.SpielerKartenNeu
        Call frmRoserl.HintersteEntladen(PlayerTurn, gintIndexVerschieben + 1)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)
End Sub
