VERSION 5.00
Begin VB.Form frmSpielBogen 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Lorum-Spielbogen"
   ClientHeight    =   13860
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   10740
   ControlBox      =   0   'False
   Icon            =   "frmSpielBogen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13860
   ScaleWidth      =   10740
   Begin VB.PictureBox Picture1 
      Height          =   20124
      Left            =   0
      ScaleHeight     =   20076
      ScaleWidth      =   10044
      TabIndex        =   1
      Top             =   0
      Width           =   10092
      Begin VB.CommandButton lblOk2 
         Caption         =   "OK"
         Height          =   492
         Left            =   5160
         TabIndex        =   196
         Top             =   18000
         Width           =   1332
      End
      Begin VB.CommandButton lblOK1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OK"
         Height          =   372
         Left            =   4560
         TabIndex        =   195
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label lblSpielerZwS_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   194
         Top             =   17400
         Width           =   372
      End
      Begin VB.Label lblSpielerEuL_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   193
         Top             =   15480
         Width           =   372
      End
      Begin VB.Label lblSpielerRK_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   192
         Top             =   15960
         Width           =   372
      End
      Begin VB.Label lblSpielerLorum_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   191
         Top             =   16440
         Width           =   372
      End
      Begin VB.Label lblSpielerSumme_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   190
         Top             =   16920
         Width           =   372
      End
      Begin VB.Label lblSpielerStiche_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   189
         Top             =   15000
         Width           =   372
      End
      Begin VB.Label lblSpielerOber_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   188
         Top             =   14520
         Width           =   372
      End
      Begin VB.Label lblSpielerRot_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   187
         Top             =   14040
         Width           =   372
      End
      Begin VB.Label lblRechtsEuL_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   186
         Top             =   15480
         Width           =   372
      End
      Begin VB.Label lblRechtsRK_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   185
         Top             =   15960
         Width           =   372
      End
      Begin VB.Label lblRechtsLorum_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   184
         Top             =   16440
         Width           =   372
      End
      Begin VB.Label lblRechtsSumme_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   183
         Top             =   16920
         Width           =   372
      End
      Begin VB.Label lblRechtsZwS_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   182
         Top             =   17400
         Width           =   372
      End
      Begin VB.Label lblRechtsStiche_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   181
         Top             =   15000
         Width           =   372
      End
      Begin VB.Label lblRechtsOber_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   180
         Top             =   14520
         Width           =   372
      End
      Begin VB.Label lblRechtsRot_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   179
         Top             =   14040
         Width           =   372
      End
      Begin VB.Label lblMitteEuL_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   178
         Top             =   15480
         Width           =   372
      End
      Begin VB.Label lblMitteRK_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   177
         Top             =   15960
         Width           =   372
      End
      Begin VB.Label lblMitteLorum_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   176
         Top             =   16440
         Width           =   372
      End
      Begin VB.Label lblMitteSumme_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   175
         Top             =   16920
         Width           =   372
      End
      Begin VB.Label lblMitteZwS_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   174
         Top             =   17400
         Width           =   372
      End
      Begin VB.Label lblMitteStiche_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   173
         Top             =   15000
         Width           =   372
      End
      Begin VB.Label lblMItteOber_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   172
         Top             =   14520
         Width           =   372
      End
      Begin VB.Label lblMitteRot_4 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   171
         Top             =   14040
         Width           =   372
      End
      Begin VB.Label lblLinksEuL_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   170
         Top             =   15480
         Width           =   372
      End
      Begin VB.Label lblLinksRK_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   169
         Top             =   15960
         Width           =   372
      End
      Begin VB.Label lblLinksLorum_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   168
         Top             =   16440
         Width           =   372
      End
      Begin VB.Label lblLinksSumme_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   167
         Top             =   16920
         Width           =   372
      End
      Begin VB.Label lblLinksZwS_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   166
         Top             =   17400
         Width           =   372
      End
      Begin VB.Label lblLinksStiche_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   165
         Top             =   15000
         Width           =   372
      End
      Begin VB.Label lblLinksOber_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   164
         Top             =   14520
         Width           =   372
      End
      Begin VB.Label lblLinksRot_4 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   163
         Top             =   14040
         Width           =   372
      End
      Begin VB.Label lblSpielerZwS_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   162
         Top             =   13560
         Width           =   372
      End
      Begin VB.Label lblMitteEuL_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   161
         Top             =   11640
         Width           =   372
      End
      Begin VB.Label lblLinksEuL_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   160
         Top             =   11640
         Width           =   372
      End
      Begin VB.Label lblSpielerEuL_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   159
         Top             =   11640
         Width           =   372
      End
      Begin VB.Label lblRechtsEuL_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   158
         Top             =   11640
         Width           =   372
      End
      Begin VB.Label lblMitteRK_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   157
         Top             =   12120
         Width           =   372
      End
      Begin VB.Label lblLinksRK_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   156
         Top             =   12120
         Width           =   372
      End
      Begin VB.Label lblSpielerRK_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   155
         Top             =   12120
         Width           =   372
      End
      Begin VB.Label lblRechtsRK_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   154
         Top             =   12120
         Width           =   372
      End
      Begin VB.Label lblMitteLorum_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   153
         Top             =   12600
         Width           =   372
      End
      Begin VB.Label lblLinksLorum_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   152
         Top             =   12600
         Width           =   372
      End
      Begin VB.Label lblSpielerLorum_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   151
         Top             =   12600
         Width           =   372
      End
      Begin VB.Label lblRechtsLorum_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   150
         Top             =   12600
         Width           =   372
      End
      Begin VB.Label lblMitteSumme_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   149
         Top             =   13080
         Width           =   372
      End
      Begin VB.Label lblLinksSumme_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   148
         Top             =   13080
         Width           =   372
      End
      Begin VB.Label lblSpielerSumme_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   147
         Top             =   13080
         Width           =   372
      End
      Begin VB.Label lblRechtsSumme_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   146
         Top             =   13080
         Width           =   372
      End
      Begin VB.Label lblMitteZwS_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   145
         Top             =   13560
         Width           =   372
      End
      Begin VB.Label lblLinksZwS_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   144
         Top             =   13560
         Width           =   372
      End
      Begin VB.Label lblRechtsZwS_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   143
         Top             =   13560
         Width           =   372
      End
      Begin VB.Label lblMitteStiche_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   142
         Top             =   11160
         Width           =   372
      End
      Begin VB.Label lblLinksStiche_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   141
         Top             =   11160
         Width           =   372
      End
      Begin VB.Label lblSpielerStiche_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   140
         Top             =   11160
         Width           =   372
      End
      Begin VB.Label lblRechtsStiche_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   139
         Top             =   11160
         Width           =   372
      End
      Begin VB.Label lblMItteOber_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   138
         Top             =   10680
         Width           =   372
      End
      Begin VB.Label lblLinksOber_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   137
         Top             =   10680
         Width           =   372
      End
      Begin VB.Label lblSpielerOber_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   136
         Top             =   10680
         Width           =   372
      End
      Begin VB.Label lblRechtsOber_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   135
         Top             =   10680
         Width           =   372
      End
      Begin VB.Label lblMitteRot_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   134
         Top             =   10200
         Width           =   372
      End
      Begin VB.Label lblLinksRot_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   133
         Top             =   10200
         Width           =   372
      End
      Begin VB.Label lblSpielerRot_3 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   132
         Top             =   10200
         Width           =   372
      End
      Begin VB.Label lblRechtsRot_3 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   131
         Top             =   10200
         Width           =   372
      End
      Begin VB.Label lblMitteZwS_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   130
         Top             =   9720
         Width           =   372
      End
      Begin VB.Label lblLinksZwS_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   129
         Top             =   9720
         Width           =   372
      End
      Begin VB.Label lblSpielerZwS_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   128
         Top             =   9720
         Width           =   372
      End
      Begin VB.Label lblRechtsZwS_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   127
         Top             =   9720
         Width           =   372
      End
      Begin VB.Label lblMitteSumme_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   126
         Top             =   9240
         Width           =   372
      End
      Begin VB.Label lblLinksSumme_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   125
         Top             =   9240
         Width           =   372
      End
      Begin VB.Label lblSpielerSumme_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   124
         Top             =   9240
         Width           =   372
      End
      Begin VB.Label lblRechtsSumme_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   123
         Top             =   9240
         Width           =   372
      End
      Begin VB.Label lblMitteLorum_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   122
         Top             =   8760
         Width           =   372
      End
      Begin VB.Label lblLinksLorum_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   121
         Top             =   8760
         Width           =   372
      End
      Begin VB.Label lblSpielerLorum_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   120
         Top             =   8760
         Width           =   372
      End
      Begin VB.Label lblRechtsLorum_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   119
         Top             =   8760
         Width           =   372
      End
      Begin VB.Label lblMitteRK_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   118
         Top             =   8280
         Width           =   372
      End
      Begin VB.Label lblLinksRK_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   117
         Top             =   8280
         Width           =   372
      End
      Begin VB.Label lblSpielerRK_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   116
         Top             =   8280
         Width           =   372
      End
      Begin VB.Label lblRechtsRK_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   115
         Top             =   8280
         Width           =   372
      End
      Begin VB.Label lblMitteEuL_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   114
         Top             =   7800
         Width           =   372
      End
      Begin VB.Label lblLinksEuL_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   113
         Top             =   7800
         Width           =   372
      End
      Begin VB.Label lblSpielerEuL_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   112
         Top             =   7800
         Width           =   372
      End
      Begin VB.Label lblRechtsEuL_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   111
         Top             =   7800
         Width           =   372
      End
      Begin VB.Label lblMitteStiche_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   110
         Top             =   7320
         Width           =   372
      End
      Begin VB.Label lblLinksStiche_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   109
         Top             =   7320
         Width           =   372
      End
      Begin VB.Label lblSpielerStiche_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   108
         Top             =   7320
         Width           =   372
      End
      Begin VB.Label lblRechtsStiche_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   107
         Top             =   7320
         Width           =   372
      End
      Begin VB.Label lblMItteOber_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   106
         Top             =   6840
         Width           =   372
      End
      Begin VB.Label lblLinksOber_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   105
         Top             =   6840
         Width           =   372
      End
      Begin VB.Label lblSpielerOber_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   104
         Top             =   6840
         Width           =   372
      End
      Begin VB.Label lblRechtsOber_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   103
         Top             =   6840
         Width           =   372
      End
      Begin VB.Label lblMitteRot_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   102
         Top             =   6360
         Width           =   372
      End
      Begin VB.Label lblLinksRot_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   101
         Top             =   6360
         Width           =   372
      End
      Begin VB.Label lblSpielerRot_2 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   100
         Top             =   6360
         Width           =   372
      End
      Begin VB.Label lblRechtsRot_2 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   99
         Top             =   6360
         Width           =   372
      End
      Begin VB.Label lblRechtsZwS_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   98
         Top             =   5880
         Width           =   372
      End
      Begin VB.Label lblSpielerZwS_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   97
         Top             =   5880
         Width           =   372
      End
      Begin VB.Label lblLinksZwS_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   96
         Top             =   5880
         Width           =   372
      End
      Begin VB.Label lblMitteZwS_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   95
         Top             =   5880
         Width           =   372
      End
      Begin VB.Label lblRechtsSumme_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   94
         Top             =   5400
         Width           =   372
      End
      Begin VB.Label lblSpielerSumme_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   93
         Top             =   5400
         Width           =   372
      End
      Begin VB.Label lblLinksSumme_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   92
         Top             =   5400
         Width           =   372
      End
      Begin VB.Label lblMitteSumme_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   91
         Top             =   5400
         Width           =   372
      End
      Begin VB.Label lblRechtsLorum_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   90
         Top             =   4920
         Width           =   372
      End
      Begin VB.Label lblSpielerLorum_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   89
         Top             =   4920
         Width           =   372
      End
      Begin VB.Label lblLinksLorum_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   88
         Top             =   4920
         Width           =   372
      End
      Begin VB.Label lblMitteLorum_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   87
         Top             =   4920
         Width           =   372
      End
      Begin VB.Label lblRechtsRK_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   86
         Top             =   4440
         Width           =   372
      End
      Begin VB.Label lblSpielerRK_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   85
         Top             =   4440
         Width           =   372
      End
      Begin VB.Label lblLinksRK_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   84
         Top             =   4440
         Width           =   372
      End
      Begin VB.Label lblMitteRK_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   83
         Top             =   4440
         Width           =   372
      End
      Begin VB.Label lblRechtsEuL_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   82
         Top             =   3960
         Width           =   372
      End
      Begin VB.Label lblSpielerEuL_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   81
         Top             =   3960
         Width           =   372
      End
      Begin VB.Label lblLinksEuL_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   80
         Top             =   3960
         Width           =   372
      End
      Begin VB.Label lblMitteEuL_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   79
         Top             =   3960
         Width           =   372
      End
      Begin VB.Label lblRechtsStiche_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   78
         Top             =   3480
         Width           =   372
      End
      Begin VB.Label lblSpielerStiche_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   77
         Top             =   3480
         Width           =   372
      End
      Begin VB.Label lblLinksStiche_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   76
         Top             =   3480
         Width           =   372
      End
      Begin VB.Label lblMitteStiche_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   75
         Top             =   3480
         Width           =   372
      End
      Begin VB.Label lblRechtsOber_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   74
         Top             =   3000
         Width           =   372
      End
      Begin VB.Label lblSpielerOber_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   73
         Top             =   3000
         Width           =   372
      End
      Begin VB.Label lblLinksOber_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   4200
         TabIndex        =   72
         Top             =   3000
         Width           =   372
      End
      Begin VB.Label lblMItteOber_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   5640
         TabIndex        =   71
         Top             =   3000
         Width           =   372
      End
      Begin VB.Label lblRechtsRot_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   6840
         TabIndex        =   70
         Top             =   2520
         Width           =   372
      End
      Begin VB.Label lblSpielerRot_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   8280
         TabIndex        =   69
         Top             =   2520
         Width           =   372
      End
      Begin VB.Label lblLinksRot_1 
         Alignment       =   2  'Zentriert
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4200
         TabIndex        =   68
         Top             =   2520
         Width           =   372
      End
      Begin VB.Label lblMitteRot_1 
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5640
         TabIndex        =   67
         Top             =   2520
         Width           =   372
      End
      Begin VB.Label Label63 
         Caption         =   "Spieler"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   7920
         TabIndex        =   66
         Top             =   2040
         Width           =   852
      End
      Begin VB.Label Label62 
         Caption         =   "Rechts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6600
         TabIndex        =   65
         Top             =   2040
         Width           =   1092
      End
      Begin VB.Label Label61 
         Caption         =   "Mitte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5280
         TabIndex        =   64
         Top             =   2040
         Width           =   972
      End
      Begin VB.Label Label60 
         Caption         =   "Links"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3840
         TabIndex        =   63
         Top             =   2040
         Width           =   1092
      End
      Begin VB.Line Line41 
         X1              =   7800
         X2              =   7800
         Y1              =   1920
         Y2              =   17760
      End
      Begin VB.Line Line40 
         X1              =   5160
         X2              =   5160
         Y1              =   1920
         Y2              =   17760
      End
      Begin VB.Line Line39 
         X1              =   6480
         X2              =   6480
         Y1              =   1920
         Y2              =   17760
      End
      Begin VB.Line Line38 
         X1              =   3720
         X2              =   3720
         Y1              =   1920
         Y2              =   17760
      End
      Begin VB.Line Line37 
         X1              =   1200
         X2              =   9000
         Y1              =   17760
         Y2              =   17760
      End
      Begin VB.Line Line36 
         X1              =   9000
         X2              =   9000
         Y1              =   1920
         Y2              =   17760
      End
      Begin VB.Line Line35 
         X1              =   1200
         X2              =   9000
         Y1              =   17280
         Y2              =   17280
      End
      Begin VB.Line Line34 
         X1              =   1200
         X2              =   9000
         Y1              =   16800
         Y2              =   16800
      End
      Begin VB.Line Line33 
         X1              =   1200
         X2              =   9000
         Y1              =   16320
         Y2              =   16320
      End
      Begin VB.Line Line32 
         X1              =   1200
         X2              =   9000
         Y1              =   15840
         Y2              =   15840
      End
      Begin VB.Line Line31 
         X1              =   1200
         X2              =   9000
         Y1              =   15360
         Y2              =   15360
      End
      Begin VB.Line Line30 
         X1              =   1200
         X2              =   9000
         Y1              =   14880
         Y2              =   14880
      End
      Begin VB.Line Line29 
         X1              =   1200
         X2              =   9000
         Y1              =   14400
         Y2              =   14400
      End
      Begin VB.Line Line28 
         X1              =   1200
         X2              =   9000
         Y1              =   13920
         Y2              =   13920
      End
      Begin VB.Line Line27 
         X1              =   1200
         X2              =   9000
         Y1              =   13440
         Y2              =   13440
      End
      Begin VB.Label Label59 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2760
         TabIndex        =   62
         Top             =   14520
         Width           =   492
      End
      Begin VB.Label Label58 
         Caption         =   "Endsumme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   2400
         TabIndex        =   61
         Top             =   17400
         Width           =   1452
      End
      Begin VB.Label Label57 
         Caption         =   "Summe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2760
         TabIndex        =   60
         Top             =   16920
         Width           =   972
      End
      Begin VB.Label Label56 
         Caption         =   "Rechts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   59
         Top             =   16440
         Width           =   732
      End
      Begin VB.Label Label55 
         Caption         =   "Mitte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   58
         Top             =   15960
         Width           =   1092
      End
      Begin VB.Label Label54 
         Caption         =   "Links"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   57
         Top             =   15480
         Width           =   852
      End
      Begin VB.Label Label53 
         Caption         =   "Spieler"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   56
         Top             =   15000
         Width           =   852
      End
      Begin VB.Label Label52 
         Caption         =   "Rechts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   55
         Top             =   14520
         Width           =   732
      End
      Begin VB.Label Label51 
         Caption         =   "Mitte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   54
         Top             =   14040
         Width           =   1092
      End
      Begin VB.Label Label50 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2760
         TabIndex        =   53
         Top             =   16440
         Width           =   492
      End
      Begin VB.Label Label49 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2760
         TabIndex        =   52
         Top             =   15960
         Width           =   612
      End
      Begin VB.Label Label48 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2760
         TabIndex        =   51
         Top             =   15480
         Width           =   492
      End
      Begin VB.Label Label47 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2748
         TabIndex        =   50
         Top             =   15000
         Width           =   492
      End
      Begin VB.Label Label46 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2748
         TabIndex        =   49
         Top             =   14040
         Width           =   492
      End
      Begin VB.Label Label45 
         Caption         =   "Zwischensumme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1800
         TabIndex        =   48
         Top             =   13560
         Width           =   1812
      End
      Begin VB.Label Label44 
         Caption         =   "Summe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2760
         TabIndex        =   47
         Top             =   13080
         Width           =   972
      End
      Begin VB.Label Label43 
         Caption         =   "Spieler"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   46
         Top             =   12600
         Width           =   732
      End
      Begin VB.Label Label42 
         Caption         =   "Rechts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   45
         Top             =   12120
         Width           =   1092
      End
      Begin VB.Label Label41 
         Caption         =   "Mitte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   44
         Top             =   11640
         Width           =   852
      End
      Begin VB.Label Label40 
         Caption         =   "Links"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   43
         Top             =   11160
         Width           =   852
      End
      Begin VB.Label Label39 
         Caption         =   "Spieler"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   42
         Top             =   10680
         Width           =   732
      End
      Begin VB.Label Label38 
         Caption         =   "Rechts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   41
         Top             =   10200
         Width           =   1092
      End
      Begin VB.Label Label37 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2760
         TabIndex        =   40
         Top             =   12600
         Width           =   492
      End
      Begin VB.Label Label36 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2760
         TabIndex        =   39
         Top             =   12120
         Width           =   612
      End
      Begin VB.Label Label35 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2760
         TabIndex        =   38
         Top             =   11640
         Width           =   492
      End
      Begin VB.Label Label34 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2748
         TabIndex        =   37
         Top             =   11160
         Width           =   492
      End
      Begin VB.Label Label33 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2760
         TabIndex        =   36
         Top             =   10680
         Width           =   492
      End
      Begin VB.Label Label32 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2748
         TabIndex        =   35
         Top             =   10200
         Width           =   492
      End
      Begin VB.Label Label30 
         Caption         =   "Zwischensumme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1800
         TabIndex        =   34
         Top             =   9720
         Width           =   1932
      End
      Begin VB.Label Label29 
         Caption         =   "Zwischensumme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1800
         TabIndex        =   33
         Top             =   5880
         Width           =   1932
      End
      Begin VB.Line Line26 
         X1              =   1200
         X2              =   9000
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line25 
         X1              =   1200
         X2              =   9000
         Y1              =   12480
         Y2              =   12480
      End
      Begin VB.Line Line24 
         X1              =   1200
         X2              =   9000
         Y1              =   12000
         Y2              =   12000
      End
      Begin VB.Line Line23 
         X1              =   1200
         X2              =   9000
         Y1              =   11520
         Y2              =   11520
      End
      Begin VB.Line Line22 
         X1              =   1200
         X2              =   9000
         Y1              =   11040
         Y2              =   11040
      End
      Begin VB.Line Line21 
         X1              =   1200
         X2              =   9000
         Y1              =   10560
         Y2              =   10560
      End
      Begin VB.Line Line20 
         X1              =   1200
         X2              =   9000
         Y1              =   10080
         Y2              =   10080
      End
      Begin VB.Line Line19 
         X1              =   1200
         X2              =   9000
         Y1              =   9600
         Y2              =   9600
      End
      Begin VB.Line Line18 
         X1              =   1200
         X2              =   9000
         Y1              =   9600
         Y2              =   9600
      End
      Begin VB.Label Label31 
         Caption         =   "Summe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2760
         TabIndex        =   32
         Top             =   9240
         Width           =   972
      End
      Begin VB.Line Line17 
         X1              =   1200
         X2              =   9000
         Y1              =   9120
         Y2              =   9120
      End
      Begin VB.Line Line16 
         X1              =   1200
         X2              =   9000
         Y1              =   12960
         Y2              =   12960
      End
      Begin VB.Label Label28 
         Caption         =   "Links"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   31
         Top             =   8760
         Width           =   732
      End
      Begin VB.Label Label27 
         Caption         =   "Spieler"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   30
         Top             =   8280
         Width           =   1092
      End
      Begin VB.Label Label26 
         Caption         =   "Rechts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   29
         Top             =   7800
         Width           =   852
      End
      Begin VB.Label Label25 
         Caption         =   "Mitte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   28
         Top             =   7320
         Width           =   852
      End
      Begin VB.Label Label24 
         Caption         =   "Links"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   27
         Top             =   6840
         Width           =   732
      End
      Begin VB.Label Label23 
         Caption         =   "Spieler"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   26
         Top             =   6360
         Width           =   1092
      End
      Begin VB.Line Line15 
         X1              =   1200
         X2              =   9000
         Y1              =   8640
         Y2              =   8640
      End
      Begin VB.Line Line14 
         X1              =   1200
         X2              =   9000
         Y1              =   8640
         Y2              =   8640
      End
      Begin VB.Line Line13 
         X1              =   1200
         X2              =   9000
         Y1              =   8160
         Y2              =   8160
      End
      Begin VB.Line Line12 
         X1              =   1200
         X2              =   9000
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Line Line11 
         X1              =   1200
         X2              =   9000
         Y1              =   7200
         Y2              =   7200
      End
      Begin VB.Line Line10 
         X1              =   1200
         X2              =   9000
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label Label22 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2760
         TabIndex        =   25
         Top             =   8760
         Width           =   492
      End
      Begin VB.Label Label21 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2748
         TabIndex        =   24
         Top             =   8280
         Width           =   612
      End
      Begin VB.Label Label20 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2748
         TabIndex        =   23
         Top             =   7800
         Width           =   492
      End
      Begin VB.Label Label19 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2748
         TabIndex        =   22
         Top             =   7320
         Width           =   492
      End
      Begin VB.Label Label18 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2748
         TabIndex        =   21
         Top             =   6840
         Width           =   492
      End
      Begin VB.Line Line9 
         X1              =   1200
         X2              =   9000
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Label Label17 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2748
         TabIndex        =   20
         Top             =   6360
         Width           =   492
      End
      Begin VB.Label Label16 
         Caption         =   "Summe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2760
         TabIndex        =   19
         Top             =   5400
         Width           =   972
      End
      Begin VB.Label Label15 
         Caption         =   "Mitte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   18
         Top             =   4920
         Width           =   852
      End
      Begin VB.Label Label14 
         Caption         =   "Links"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   17
         Top             =   4440
         Width           =   732
      End
      Begin VB.Label Label13 
         Caption         =   "Spieler"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   16
         Top             =   3960
         Width           =   1092
      End
      Begin VB.Label Label12 
         Caption         =   "Rechts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   15
         Top             =   3480
         Width           =   852
      End
      Begin VB.Label Label11 
         Caption         =   "Mitte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   14
         Top             =   3000
         Width           =   852
      End
      Begin VB.Label Label10 
         Caption         =   "Links"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1200
         TabIndex        =   13
         Top             =   2520
         Width           =   732
      End
      Begin VB.Line Line8 
         X1              =   1200
         X2              =   9000
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line7 
         X1              =   1200
         X2              =   9000
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line6 
         X1              =   1200
         X2              =   9000
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line5 
         X1              =   1200
         X2              =   9000
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line4 
         X1              =   1200
         X2              =   9000
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line3 
         X1              =   1200
         X2              =   9000
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label9 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2760
         TabIndex        =   12
         Top             =   4920
         Width           =   492
      End
      Begin VB.Label Label8 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2748
         TabIndex        =   11
         Top             =   4440
         Width           =   612
      End
      Begin VB.Label Label7 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2748
         TabIndex        =   10
         Top             =   3960
         Width           =   492
      End
      Begin VB.Label Label6 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2750
         TabIndex        =   9
         Top             =   3480
         Width           =   492
      End
      Begin VB.Label Label5 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2750
         TabIndex        =   8
         Top             =   3000
         Width           =   492
      End
      Begin VB.Line Line2 
         X1              =   1200
         X2              =   9000
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label4 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2750
         TabIndex        =   7
         Top             =   2520
         Width           =   492
      End
      Begin VB.Label Label3 
         Caption         =   "Spiel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2760
         TabIndex        =   6
         Top             =   2040
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "Geber"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1200
         TabIndex        =   5
         Top             =   2040
         Width           =   1452
      End
      Begin VB.Line Line1 
         X1              =   1200
         X2              =   9000
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Caption         =   "Lorum-Spielbogen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2160
         TabIndex        =   4
         Top             =   1200
         Width           =   5652
      End
      Begin VB.Label lblDatum 
         Height          =   372
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   2892
      End
      Begin VB.Label Datum 
         Caption         =   "Datum:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3135
      LargeChange     =   10
      Left            =   10440
      Max             =   30000
      TabIndex        =   0
      Top             =   0
      Width           =   252
   End
End
Attribute VB_Name = "frmSpielBogen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Dim n As Long, picHeight As Long
    
    Me.ScaleMode = vbPixels
    Me.WindowState = vbNormal    'vbMaximized
    'Me.Show ' force the resize now
    'Me.Caption = Me.ScaleHeight
    picHeight = Me.ScaleHeight  ' the Form's ScaleHeight
    With Picture1
        .BorderStyle = vbBSNone
        .ScaleMode = vbPixels
        .Width = Me.ScaleWidth - VScroll1.Width
        .Height = Me.ScaleHeight * picHeight
        '.Move 0, 0, Me.ScaleWidth - VScroll1.Width, Me.ScaleHeight * 2
        .Move 0, 0, Me.ScaleWidth - VScroll1.Width, Me.ScaleHeight + 500
    End With
    With VScroll1
        .Move Picture1.Width, 0, VScroll1.Width, Me.ScaleHeight
        .Min = 0
        .Max = Me.ScaleHeight - Picture1.Height
        .Value = 0
        .SmallChange = 10
        .LargeChange = 100
        .TabStop = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If blnUnloadFrmSpielBogen = False Then
'        Cancel = True
'        Me.Hide
'    End If
End Sub

Private Sub lblOK1_Click()
    Me.Hide                                                             'Gerbing 19.12.2017
End Sub

Private Sub lblOk2_Click()
    Me.Hide                                                             'Gerbing 19.12.2017
End Sub

Private Sub VScroll1_Change()
    Picture1.Top = (VScroll1.Value)
End Sub

Private Sub VScroll1_Scroll()
    Picture1.Top = (VScroll1.Value)
End Sub

