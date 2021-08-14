VERSION 5.00
Begin VB.Form frmFontSize 
   BackColor       =   &H00C0C0C0&
   Caption         =   "FontSize"
   ClientHeight    =   2724
   ClientLeft      =   48
   ClientTop       =   456
   ClientWidth     =   4824
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2724
   ScaleWidth      =   4824
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Height          =   1812
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4572
      Begin VB.OptionButton OptGross 
         Caption         =   "groß"
         Height          =   372
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2892
      End
      Begin VB.OptionButton OptMittel 
         Caption         =   "Mittel"
         Height          =   372
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2652
      End
      Begin VB.OptionButton OptKlein 
         Caption         =   "klein"
         Height          =   372
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2532
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frmFontSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call AnpassenNutzerWunsch(Me)                                       'Gerbing 11.03.2017
    Me.Caption = LoadResString(2536 + Sprache)          'Schriftgröße
    btnOK.Caption = LoadResString(3001 + Sprache)       'OK
    OptKlein.Caption = LoadResString(1845 + Sprache) 'klein             'Gerbing 11.03.2017
    OptMittel.Caption = LoadResString(1846 + Sprache) 'mittel           'Gerbing 11.03.2017
    OptGross.Caption = LoadResString(1847 + Sprache) 'gross             'Gerbing 11.03.2017
    
    If PublicCheckForDPI = "1" Then
        OptKlein.Value = True
    End If
    If PublicCheckForDPI = "2" Then
        OptMittel.Value = True
    End If
    If PublicCheckForDPI = "3" Then
        OptGross.Value = True
    End If
End Sub
        
Private Sub OptGross_Click()                                            'Gerbing 11.03.2017
    PublicCheckForDPI = "3"
    WriteDPI ("3")
End Sub

Private Sub OptKlein_Click()                                            'Gerbing 11.03.2017
    PublicCheckForDPI = "1"
    WriteDPI ("1")
End Sub

Private Sub OptMittel_Click()                                           'Gerbing 11.03.2017
    PublicCheckForDPI = "2"
    WriteDPI ("2")
End Sub
