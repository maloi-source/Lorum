VERSION 5.00
Object = "{54A43060-C747-4026-8E08-2A801095C838}#1.0#0"; "roserl.ocx"
Begin VB.Form frmRoserl 
   BackColor       =   &H00008000&
   Caption         =   "Roserl"
   ClientHeight    =   12024
   ClientLeft      =   60
   ClientTop       =   -60
   ClientWidth     =   18972
   ForeColor       =   &H000000FF&
   Icon            =   "frmRoserl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1002
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1581
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.Timer TimerEinsammlerSpieler 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10440
      Top             =   1320
   End
   Begin VB.Timer TimerEinsammlerRechts 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9840
      Top             =   1320
   End
   Begin VB.Timer TimerEinsammlerMitte 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9240
      Top             =   1320
   End
   Begin VB.Timer TimerEinsammlerLinks 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8640
      Top             =   1320
   End
   Begin VB.Timer TimerRechtsAnimation 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10440
      Top             =   360
   End
   Begin VB.Timer TimerMitteAnimation 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9840
      Top             =   360
   End
   Begin VB.Timer TimerLinksAnimation 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9240
      Top             =   360
   End
   Begin VB.Timer TimerSpielerAnimation 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8640
      Top             =   360
   End
   Begin VB.Timer TimerWeiterspielenNachAss 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   240
   End
   Begin VB.Timer SchlafTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1680
      Top             =   240
   End
   Begin RoserlOcx.Roserl cdRoserlEichelB 
      Height          =   1860
      Index           =   0
      Left            =   6600
      TabIndex        =   31
      Top             =   9000
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlEichelA 
      Height          =   1860
      Index           =   0
      Left            =   3480
      TabIndex        =   30
      Top             =   9000
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlGr�nB 
      Height          =   1860
      Index           =   0
      Left            =   6600
      TabIndex        =   29
      Top             =   6840
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlGr�nA 
      Height          =   1860
      Index           =   0
      Left            =   3480
      TabIndex        =   28
      Top             =   6840
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlRotB 
      Height          =   1860
      Index           =   0
      Left            =   6600
      TabIndex        =   27
      Top             =   4680
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlRotA 
      Height          =   1860
      Index           =   0
      Left            =   3480
      TabIndex        =   26
      Top             =   4680
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlSchellB 
      Height          =   1860
      Index           =   0
      Left            =   6600
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlSchellA 
      Height          =   1860
      Index           =   0
      Left            =   3480
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlEichelAss 
      Height          =   1860
      Index           =   0
      Left            =   17760
      TabIndex        =   23
      Top             =   9000
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlGr�nAss 
      Height          =   1860
      Index           =   0
      Left            =   17760
      TabIndex        =   22
      Top             =   6840
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlRotAss 
      Height          =   1860
      Index           =   0
      Left            =   17760
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRoserlSchellAss 
      Height          =   1860
      Index           =   0
      Left            =   17760
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl SpielerGespielt 
      Height          =   732
      Left            =   7560
      TabIndex        =   18
      Top             =   11760
      Visible         =   0   'False
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   1291
      CardType        =   3
   End
   Begin RoserlOcx.Roserl RechtsGespielt 
      Height          =   732
      Left            =   14400
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   1291
      CardType        =   3
   End
   Begin RoserlOcx.Roserl MitteGespielt 
      Height          =   732
      Left            =   7440
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   1291
      CardType        =   3
   End
   Begin RoserlOcx.Roserl LinksGespielt 
      Height          =   732
      Left            =   720
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   1291
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdSpieler 
      Height          =   1860
      Index           =   0
      Left            =   5880
      TabIndex        =   14
      Top             =   11040
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdRechts 
      Height          =   1860
      Index           =   0
      Left            =   21240
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdMitte 
      Height          =   1860
      Index           =   0
      Left            =   5880
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin RoserlOcx.Roserl cdLinks 
      Height          =   1860
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   3281
      CardType        =   3
   End
   Begin VB.TextBox txtFont 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   2
      Text            =   "txtFont"
      Top             =   11040
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblRunde 
      Alignment       =   2  'Zentriert
      Caption         =   "1. Runde"
      Height          =   252
      Left            =   5400
      TabIndex        =   19
      Top             =   6240
      Width           =   2892
   End
   Begin VB.Label lblRoserl 
      Alignment       =   2  'Zentriert
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   5
      Left            =   7800
      TabIndex        =   10
      Top             =   5760
      Width           =   492
   End
   Begin VB.Label lblRoserl 
      Alignment       =   2  'Zentriert
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   4
      Left            =   7320
      TabIndex        =   9
      Top             =   5760
      Width           =   492
   End
   Begin VB.Label lblRoserl 
      Alignment       =   2  'Zentriert
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   3
      Left            =   6840
      TabIndex        =   8
      Top             =   5760
      Width           =   492
   End
   Begin VB.Label lblRoserl 
      Alignment       =   2  'Zentriert
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   2
      Left            =   6360
      TabIndex        =   7
      Top             =   5760
      Width           =   492
   End
   Begin VB.Label lblRoserl 
      Alignment       =   2  'Zentriert
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   5880
      TabIndex        =   6
      Top             =   5760
      Width           =   492
   End
   Begin VB.Label lblRoserl 
      Alignment       =   2  'Zentriert
      BackColor       =   &H000000FF&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   5400
      TabIndex        =   5
      Top             =   5760
      Width           =   492
   End
   Begin VB.Label lblRechts 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Rechts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   14400
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   696
   End
   Begin VB.Label lblLinks 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Links"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   528
   End
   Begin VB.Label lblSpieler 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Spieler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6840
      TabIndex        =   1
      Top             =   11280
      Visible         =   0   'False
      Width           =   696
   End
   Begin VB.Label lblMitte 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mitte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6840
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   504
   End
   Begin VB.Menu mnuSpielbogen 
      Caption         =   "Spiel-Abrechnung zeigen"
   End
   Begin VB.Menu mnuLetztenStich 
      Caption         =   "Letzten Stich zeigen"
   End
   Begin VB.Menu mnuBestwerte 
      Caption         =   "Bestwerte zeigen"
   End
   Begin VB.Menu mnuFontSize 
      Caption         =   "Schriftgr��e"
   End
   Begin VB.Menu mnuVersion 
      Caption         =   "Version"
   End
   Begin VB.Menu mnuSpielregeln 
      Caption         =   "Spielregeln"
   End
   Begin VB.Menu mnuBeenden 
      Caption         =   "Beenden"
   End
End
Attribute VB_Name = "frmRoserl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Bei Eingabe von maloi deckt der Computer seine Karten auf
'Bei dr�cken der Tasten Strg+N kann das aktuelle Spiel mit den gleichen Karten noch einmal gespielt werden
'Bei dr�cken der Tasten Strg+S wird der Lorum-Spielbogen angezeigt
'Bei dr�cken der Tasten Strg+E kommt eine MsgBox ob das Spiel abgebrochen werden soll

'30.06.2017 Bei Spielart=Lorum in der Prozedur HintersteEntladen trat der Fehler auf, dass 'UBound(...Cards) <> cd....Count' ist
'           Die Ursache des Fehlers habe ich nicht gefunden, aber eine Methode um den Fehler zu korrigieren
'           Ich muss 'ReDim Preserve xxxCards(1 To cdxxx.Count)' benutzen
'           L�sung: ich lasse s�mtliche Bezugnahmen auf ComputerCards weg
'                   und ich benutze f�r Spieart=Lorum eine eigene Prozedur 'LorumEntladen'
'
'30.06.2017 Prototypfehler Es war immer bei 8 Karten Ubound(...Cards) = 9
'
'02.07.2017 seltsamer Effekt Absturz bei LorumAssAblegen
'           vb6.exe st�rzt ab "Visual Basic funktioniert nicht mehr"
'           Ursache waren unterschiedlicher Funktionsaufruf: rc = LorumAssAblegen(SpielerCards)
'           und Funktions-Declaration: Private Function LorumAssAblegen()
'
'06.07.2017 Verbesserung: Strategie beim Ober-Spiel beim Abwerfen und beim Zugeben
'06.07.2017 Verbesserung: Bei LorumEineKarteHinlegen nicht einfach die erste nehmen sondern untersuchen ob wenigstens eine Teil-Kartenfolge entstehen kann
'09.07.2017 Verbesserung: Bei den ersten zwei Stichen bei spielart = Rot kann die h�chste Karte zugegeben werden
'11.07.2017 Verbesserung: bei Spielart=Rot bei den ersten zwei Stichen die h�chste Nicht-Rot ausspielen
'10.07.2017 Verbesserung: Zeitmessungen haben ergeben - es dauert ewig lange bis zu Spielbeginn alle Karten aufgelegt sind,
'           genauso bei Form_Resize
'           Ursache: Es dauert ewig lange ein Bild aus der Ressourcendatei RoserlBlatt.RES in sein Skatblatt-Image zu laden 16 Millisek
'           Schuld war eine Function LoadImageFromRes ohne Benutzung von GDI+ im Roserl.ocx
'           L�sung: ich habe www.ActiveVB.de eine Function LoadImageFromCustomRes gefunden da dauert es nur noch 1 Millisek
'11.07.2017 Verbesserung: Lorum kann mit aufgedeckten Karten gespielt werden
'12.07.2017 Verbesserung: Bei Rot kann es theoretisch passieren, dass der Computer bei den ersten zwei Stichen garnicht anders kann
'           als Rot zu spielen
'14.07.2017 Warum dauert die Schleife zum Bewegen der Ausgespielten Karte in verschiedenen Betriebssystemen verschieden lange
'           Es liegt am xxxGespielt.Move ich lasse Move nicht 100 mal dran kommen sondern nur 25 mal oder 50 mal
'29.07.2017 Verbesserung Letzten Stich anzeigen
'31.07.2017 Verbesserung Bestwerte Zeigen benutzt Access-Datenbank RoserlBestwerte.mdb Tabelle Bestwerte
'02.08.2017 R�ckg�ngig machen von 'Lorumkarten aufgedeckt'
'13.08.2017 Fehlerkorrektur
'           Zustand: Wenn der Spieler dran ist, kann er mehr als eine Karte spielen, sooft wie er klickt
'           L�sung: Schalter blnEingesammelt
'13.08.2017 Verbesserung
'           Zustand: In einer MsgBox kann man die Schriftgr��e nicht ver�ndern
'           L�sung: Jede MsgBox wird durch eine frmMsgBox ersetzt, dort kann die Schriftgr��e klein mittel gro� sein
'                   msg wird durch gstrMsg ersetzt
'                   MsgBox wird durch frmMsgBox.Show 1 ersetzt
'                   Die Anzahl Zeilen in gstrMsg wird gez�hlt - Multiplizieren mit dem Wert f�r klein/mittel/gross=Form.Height
'                   Die breiteste Zeile in gstrMsg ermitteln - Multiplizieren mit dem Wert f�r klein/mittel/gross=Form.Width
'                   Die Form in Fenstermitte positionieren - Form.Height \ 2 und Form.Width \ 2
'                   Form sichtbar machen
'                   Die Auswahl von klein mittel gro� geschieht in frmFontSize
'                   frmFontSize erreicht man �ber das Men�
'                   Roserl.ini speichert die gew�hlte Schriftgr��e
'15.08.2017 Verbesserung
'           Zustand: Es ist schwer erkennbar wer ausgespielt hat
'           L�sung:  Der Spielername des Ausspielers wird fett und gelb dargestellt
'15.08.2017 Fehlerkorrektur
'           Zustand: Es kommt die MsgBox 'Lorum wurde von ... gewonnen', aber es liegt noch eine Karte auf dem Spielfeld
'           L�sung: Alle Karten des Gewinners auf visible = False setzen
'15.08.2017 Fehlerkorrektur
'           Zustand: Sch�nheitsfehler, Wenn der Spieler die letzte Karte ablegt, bewegt sich diese, aber es liegt noch eine drunter
'           L�sung: Bei Beginn der Bewegungsanimation(EineKarteAblegen) m�ssen alle Karten cdSpieler(i).Visible = False sein
'16.08.2017 Fehlerkorrektur
'           Zustand: Wenn bei Spielart=Lorum die letzte Karte abgelegt wird, zeigt manchmal die frmMsgBox falsche Farbe und falschen Wert
'           L�sung: Bei der letzten Karte bei Spielart=Lorum brauche ich nicht zu sagen, welche Karte das ist
'                   es gen�gt strAbleger & " legt jetzt seine letzte Karte ab"
'17.08.2017 Verbesserung
'           Nach 4 Durchg�ngen kommt die Frage 'Willst Du Weiter spielen?'
'22.08.2017 Verbesserung
'           Zustand: Vor dem Sichtbarwerden der Roserl-Karten sieht man kurz alle Karten in halber Groesse
'           L�sung: frmRoserl_Resize zu Beginn Me.Visible = False und am Ende Me.Visible = True
'22.08.2017 Verbesserung
'           Zustand: Wenn Spieler die letzte Karte hinlegt, flackert links davon manchmal noch eine Karte auf
'           L�sung: 'SpielerGespielt.Visible = True' darf erst kommen wenn die Bewegungsanimation bereits begonnen hat
'03.09.2017 Verbesserung
'           Zustand: Wenn nach 4 Spielen weiter gespielt werden soll, wurde der Spielbogen nicht gel�scht
'           L�sung: Call SpielbogenNull war vergessen worden
'04.09.2017 Fehlerkorrektur
'           Zustand: Es war Spielart = Rot. Es war der zweite Stich. Links war Ausspieler und spielt kein Rot. Rechts hat diese Farbe
'                   nicht und kann abwerfen. Rechts wirft Rot Ass ab, das ist beim ersten und zweiten Stich verboten
'           L�sung: In der Funktion NachbehandlungAbwerfen muss ich verhindern dass eine einzelne Rot-Karte ermittelt wird, sofern
'                   es sich um Spielart = Rot handelt und es ist der erste oder zweite Stich
'05.09.2017 Verbesserung
'           Zustand: Bei den Spielern Links Mitte Rechts beginnt und endet die Bewegungsanimation jedesmal an einer anderen Stelle.
'                   Das sieht unsauber aus.
'           L�sung: Der Animations-Startpunkt richtet sich an der Fenstermitte aus
'05.09.2017 Fehlerkorrektur
'           Zustand: Nach Strg+N wird zwar das Spiel mit den gleichen Karten wiederholt, aber der Ausspieler wandert weiter
'           L�sung: ich habe nicht genug Werte auf Anfangszustand zur�ckgesetzt
'06.09.2017 Sch�nheitskorrektur
'           Zustand: Bei Spielbeginn f�ngt der Computer sofort an Karten auszuspielen.
'           L�sung: Besser ist die Aufforderung "Du musst zum Weiterspielen auf den roten Buchstabe klicken"
'08.09.2017 Fehlerkorrektur
'           Wenn der Spieler in seinen Karten herumklickt obwohl er nicht dran ist, kann es passieren dass eine Karte von Spieler verschwindet.
'           Am Ende haben dann Links Mitte Rechts noch eine Karte, aber Spieler nicht.
'           L�sung: If Ausspieler <> Spieler And blnRechtsWarDran = False Then Exit Sub
'10.09.2017 Versuchte Verbesserung mit MultiMonitor-System
'           Versuch eine MDIForm1 als Mutter aller MDI-Childs zu benutzen wurde verworfen weil:
'           Form.StartUpPosition=1 (Fenstermitte) bleibt wirkungslos, ich muss es selbst machen
'           Alle MDI-Child-Formen m�ssen nichtmodal sein. Da kann ich nicht warten bis 'Der Stich geht an ..." auf OK geklickt wird.
'               Nach dem Ausgeben der Nachricht arbeitet das Programm bereits weiter. Klimmz�ge sind n�tig.
'           Das Men� muss in die MDIForm1 verlegt werden
'           Ich muss Klimmz�ge mit frmRoserl_Resize machen
'---------------------------------------------------------------------------------------------------------------------------
'Version 4
'11.09.2017 Erfolgreiche Verbesserung mit MultiMonitor-System
'           damit frmMsgBox seine Ausgaben auf dem Zweitmonitor erzeugt
'           ich hatte keinen echten Zweitmonitor sondern einen zweiten Rechner und dessen Monitor als Zweitmonitor benutzt
'           das schafft man mit der Software SpaceDesk (datronicsoft) ich muss Windows10 benutzen
'           Klassenmodul clsScreens benutzt
'           Intuitive �nderung von Me.top durch addition von 5000, sonst sa� die Schrift zu weit oben
'---------------------------------------------------------------------------------------------------------------------------
'Version 5
'12.09.2017 Verbesserung
'           Version ins Men� eintragen
'12.09.2017 Verbesserung
'           Spiel beenden, wenn alle Ober gespielt sind oder alle Roten gespielt sind
'13.09.2017 Verbesserung
'           auf Hinweis von Hans-J�rgen Richter gibt es ab sofort eine portable Version
'           Zu diesem Zweck muss ich auf RoserlBestwerte.mdb verzichten und statt dessen eine RoserlBestwerte.log benutzen
'           und anstelle des Msdatgrid ein MSFlexGrid benutzen
'18.09.2017 Fehlerkorrektur
'           Zustand: Manchmal ist in der Roserl-Ergebnisliste im ersten Runde alles Null, Eintr�ge beginnen erst ab zweiten Runde
'           L�sung: 'Rundez�hler erh�hen war an der falschen Stelle durchgef�hrt worden
'---------------------------------------------------------------------------------------------------------------------------
'Version 6
'19.09.2017 Fehlerkorrektur
'           Bei Spielart = Lorum h�ngt sich das Programm auf, sobald das Taskleisten-Icon geklickt wird.
'           Es geht in eine endlose Form_Resize-Schleife
'           L�sung: Ich finde keine bessere L�sung als bei Spielart = Lorum kommt Form_Resize nicht dran
'                   Wieder r�ckg�ngig gemacht
'                   Rorm_Resize bei Lorum geht jetzt
'19.09.2017 Fehlerkorrektur
'           Zustand: Der Wert f�r Stunde ist 2 Stunden zu klein bei Eintragen Bestwerte
'           L�sung: Ich addiere 2 Stunden au�er wenn es die 24 �berschreiten w�rde, da addiere ich 1 zum Tag
'20.09.2017 Verbesserung
'           Zustand: Manchmal ist das Spielfeld zu klein. zB bei Resolution kleiner 1920 x 1200 oder bei Lorum
'           L�sung: Ich ordne die LinksKarten und die RechtsKarten entlang des Spielfeldrandes �bereinander an
'                   Durch die neue Anordnung �berdecken manchmal die RechtsKarten einen Teil der Spielerkarten. Dagegen
'                   hilft ZOrder
'20.09.2017 Verbesserung
'           Zustand: Nach jedem Stich kommt MsgBox "Der Stich geht an..." Mich nervt, dass ich laufend auf OK klicken muss
'           L�sung: Es wird eine Bewegungsanimation der gespielten Karten zum Einsammler hin gemacht
'20.09.2017 Verbesserung
'           Schlaftimer.Enabled = True wird auch gesetzt wenn RechtsPlay dran war
'22.09.2017 Fehlerkorrektur
'           Wenn Schlaftimer.Enabled = True gesetzt wird kann es passieren dass gerade eine MsgBox beantwortet werden soll und jetzt
'           Kommt schon wieder eine neue MsgBox mit "Schlaf nicht..." das f�hrt zu Fehler 400 eine Modale Form kann nicht geladen
'           werden, wenn sie schon geladen ist
'---------------------------------------------------------------------------------------------------------------------------
'Version 7
'25.09.2017 Verbesserung
'           Zustand: Beim Lorum-Spiel nervt mich der Zwang laufend auf OK zu klicken
'           L�sung: Es wird eine Bewegungsanimation der gespielten Karte zum Zielpunkt hin gemacht
'26.09.2017 Fehlerkorrektur
'           Zustand: Es kommt an falscher Stelle 'Schlaf nicht'
'           L�sung: Besseres ein/ausschalten
'26.09.2017 Verbesserung
'           Zustand: G�nter hat kritisiert, dass das Programm des Spielers eventuelle Dummheit korrigiert und Abw�rts-Ablegekarten
'                   eigenm�chtig ablegt. Wenn der Spieler zu bl�d ist, soll ihm das Programm nicht helfen
'                   Genauso, wenn der Spieler zwei Asse ablegen kann, greift manchmal das Programm ein und legt ein
'                   anderes Ass ab als das angeklickte
'           L�sung: Der Spieler hat ab sofort die volle Hoheit
'27.09.2017 Zustand: Die bei Lorum animierte Karte versteckt sich hinter den anderen
'           L�sung: Bei MoveKarteLorum wird die animierte Karte mit ZOrder nach oben gelegt
'                   Das ZOrder wirkt offenbar weiter auch bei den n�chsten Spielen, da �berlagert zB cdMitte(i) die Karte MitteGespielt
'                   Ich muss also auch LinksGespielt MitteGespielt RechtsGespielt mit ZOrder nach oben bringen
'28.09.2017 Zustand: Lorum - Bei mehr als einem Ass wird das n�chste kurz mal aufgedeckt, auch bei Beginn einer Kartenfolge
'           L�sung: erst bei MoveKarteLorum wird die zu bewegende Karte aufgedeckt
'28.09.2017 Zustand: Beim 5.Runde m�sste 1.Runde dort stehen
'           L�sung: Das war vergessen worden. Das hat bisher noch niemand gemerkt
'28.09.2017 Zustand: Lorum - Die Position der ersten Ablegekarte liegt manchmal zu weit links. das passiert nie beim ersten Lorum-Spiel
'           L�sung: Ich muss mir die Position aller Lorum-Ablegekarten merken und vor jedem Lorum-Spiel wiederherstellen
'03.10.2017 Zustand: Bestwerte - Wenn zwei Spieler die gleiche niedrige Punktzahl haben ist strPunkte und strSpieler leer
'                   nur Datum steht drin
'           L�sung: < ge�ndert in <=
'03.10.2017 Zustand: Laufend wechselt die Geschwindigkeit bei der Bewegungsanimation
'           L�sung: Nicht mehr Links und Spieler werden gemessen, sondern nur noch Spieler
'----------------------------------------------------------------------------------------------------------------------------
'Version 8
'04.10.2017 Verbesserung
'           Zustand: Spielart = Lorum Es kommt eine Msgbox
'                   gstrMsg = "Der n�chste Ausspieler ist " & strNextPlayer & vbNewLine
'                   gstrMsg = gstrMsg & "Du musst auf die Karten des Ausspielers Klicken" & vbNewLine
'           Das Programm soll selber auf die Karten des Ausspielers klicken, damit wird es noch etwas fl�ssiger
'           L�sung: zB Call cdLinks_Click(0)
'05.10.2017 Fehlerkorrektur:
'           Zustand: Das Programm soll eine Msgbox bringen "Roserl verlangt eine Bildschirmaufl�sung von mindestens 1920x1060."
'                   Stattdessen kommt Laufzeitfehler '91' "Objektvariable oder With-Blockvariable nicht festgelegt"
'           L�sung: Vertauschen von
'                   Set Mons = New clsScreens                                           'Gerbing 05.10.2017
'                   Call SetStandardWerteLorum
'                   Set Mons = ... muss vor SetStandardWerteLorum kommen sonst ist Mons noch nicht definiert
'05.10.2017 Fehlerkorrektur
'           Zustand: Bei Men�-Auswahl Schriftgr��e kommt Fehler '326' "Ressource mit dem Bezeichner '2536' nicht gefunden"
'           L�sung: Ich hatte Language.res als Projekt-Bestandteil und habe es entfernt
'                   Dann habe ich vergessen, dass frmFontsize mit LoadResString arbeitet
'                   Ich muss LoadResString entfernen
'05.10.2017 Zustand: LoadImageFromCustomRes bringt Fehler '326' Resource mit dem Bezeichner '409' oder '309 nicht gefunden
'           L�sung: LorumStartWert = LorumStartWert + 1 darf nicht den Wert 9 erreichen
'----------------------------------------------------------------------------------------------------------------------------
'Version 9
'13.10.2017 Verbesserung der Intelligenz
'           Zustand: Bisher habe ich nicht beachtet, dass bei Spielart=Rot ab der dritten Karte vorrangig Rot abgeworfen werden muss
'               Forderung 12.10.2017 Uwe Riedel: Unterschiedliche Raffinesse der 3 Mitspieler
'               1.links mitte rechts sollten doch etwas sch�rfer spielen:
'               - die drei werfen im Rot-Spiel kein bzw. kaum rot ab, wenn sie die gespielte Farbe nicht haben.
'           L�sung: Zus�tzliche Abfragen, Zuerst die h�chste Rot, dann h�chste andere Farben abwerfen
'           Zustand: Ich habe zwar vorgesehen, dass der Computer keine Karte ausspielen soll, die nur er noch selber hat,
'                   aber ich habe es falsch programmiert
'                   Forderung 12.10.2017 Uwe Riedel- rechts z.B. spielte aus Schell 7 (die 9 hat er noch selber),
'                   obwohl alle anderen 6 Schellen schon gespielt sind.Er bekommt m�chtig Rot-Karten.
'                   If ComputerCards(I).cGespieltVon = 1 Then 'Falsch
'           L�sung:
'                   If ComputerCards(I).cGespieltVon <> 1 Then 'Richtig
'                   und wenn schon 8 Karten gez�hlt wurden, dann muss ich diese Farbe beim Weitersuchen �berspringen
'14.10.2017 Verbesserung
'           Zustand: Uwe Riedel: bei ROSER wird immer nach dem Dr�cken des roten Buchstaben ausgespielt, aber bei L nicht.
'                   Und ich habe mir nicht gemerkt, bzw. ich lese es vorher garnicht wer dran ist.
'                   Nun mu� ich erst testen wer dran ist. Also bitte auch bei rotes L dr�cken ausspielen lassen.
'           L�sung: weglassen "Der Lorum-Ausspieler ist " & strAusspieler
'                   Nur bei Spielart=Rot kommt "Du musst zum Weiterspielen auf den roten Buchstabe klicken" sonst weglassen
'                   Die erste Lorum-Karte spielt der Computer aus
'14.10.2017 Verbesserung der Intelligenz
'           Zustand: Spielart = Rot: Links ist Ausspieler f�r den 3.Stich und spielt sein h�chste Karte aus,
'                   er h�tte seine niedrigste suchen m�ssen
'           L�sung:         If Spielart = Rot And AKIndex < 9 Then                     'falsch
'                           If Spielart = Rot And AKIndex < 8 Then                     'richtig
'                   er hat gedacht er ist noch beim 2.Stich, da w�re es korrekt gewesen
'15.10.2017 Verbesserung der Intelligenz
'           Zustand: Bei Rot und Stiche soll NachbehandlungMitspieler herausfinden ob die vorgeschlagene Karte NiedrigstePosition lieber nicht
'                   gespielt werden soll. Das soll nur passieren, beim erstenmal Zugeben auf diese Farbe
'                   Es passiert auch beim zweiten Zugeben und das ist falsch
'           L�sung: Ich hatte vergessen auch bei den ersten zwei Stichen bei Spielart=Rot/Stiche die Schalter bln...ZEM
'                   ZEM = Zum Ersten Mal auf False zu setzen
'15.10.2017 Verbesserung der Intelligenz
'           Zustand: Bisher wird bei Spielart = Ober und Spielart = RoterK�nig keine NachbehandlungAusspieler gemacht, ob vielleicht
'                   eine Farbe ausgespielt werden soll, die keiner mehr hat
'           L�sung: Aufruf von NachbehandlungAusspieler
'15.10.2017 Verbesserung der Intelligenz
'           Zustand: Es ist sinnlos eine Farbe abzuwerfen, die niemand mehr ausspielen kann, au�er ich habe keine andere
'           L�sung: Zus�tzliche Kontrolle in NachbehandlungAbwerfen -> HatKeinerMehr
'15.10.2017 Verbesserung der Intelligenz
'           Zustand: NachbehandlungAbwerfen will anstelle von DieH�chsteAbwerfen manchmal besser DieEinzigeAbwerfen
'                   wenn es nicht gerade die 7 ist
'           L�sung: wenn es nicht gerade die 7 oder 8 oder 9 ist
'16.10.2017 Verbesserung der Intelligenz
'           Bei Rot oder Stiche
'           Wenn ich der 4.Zugeber(der letzte eines Stiches) bin, muss ich versuchen
'           geradeso unterhalb der h�chsten Karte zu bleiben, die auf dem Tisch liegt(3 liegen auf dem Tisch), nicht einfach die niedrigste
'           zugeben -> VorbehandlungRotOderStiche
'16.10.2017 Verbesserung der Intelligenz
'           Wenn der Ober noch nicht gespielt ist und nicht auf dem Tisch liegt
'           und wenn ich der 4.Zugeber(der letzte eines Stiches) bin, kann ich K�nig oder Ass zugeben
'           nicht einfach die niedrigste zugeben, sondern die h�chste suchen -> VorbehandlungOber
'16.10.2017 Verbesserung der Intelligenz
'           Wenn der RoterK�nig noch nicht gespielt ist und nicht auf dem Tisch liegt
'           und wenn ich der 4.Zugeber(der letzte eines Stiches) bin, kann ich Ass zugeben
'           nicht einfach die niedrigste zugeben, sondern die h�chste suchen -> VorbehandlungRoterK�nig
'----------------------------------------------------------------------------------------------------------------------------
'Version 10
'20.10.2017 Verbesserung des fl�ssigen Spiels
'           Zustand: Bei Spielart = Lorum muss nach jedem Ass vom Spieler geklickt werden. Das soll das Programm �bernehmen
'           L�sung: Ich benutze den Timer TimerWeiterspielenNachAss
'23.10.2017 Verbesserung:
'           Zustand: Das Beenden �ber das Schlie�kreuz funktioniert nur selten
'           L�sung: End benutzen, damit wird es auch nicht besser
'02.11.2017 Fehlerkorrektur
'           Zustand: EintragenBestwerte: Nach Ende der Sommerzeit ist die Uhrzeit eine Stunde h�her als wirklich
'           L�sung: Function GMTDiff()
'04.11.2017 Verbesserung der Intelligenz
'           Zustand: Bei Spielart = Rot fehlt es immer noch an Intelligenz beim Ausspielen
'           L�sung: Zuerst nach Rot9 oder Rot8 oder Rot7 suchen
'08.11.2017 Verbesserung der Fl�ssigkeit
'           Versuche gegen das Flackern bei der Bewegungsanimation
'           MaxDoEvents darf nicht >=2 sein
'09.11.2017 Verbesserung der Fl�ssigkeit
'           Weitere Versuche gegen das Flackern bei der Bewegungsanimation
'           Bewegungsanimation nicht mit unterschiedlich vielen DoEvents und unterschiedlich vielen Schleifendurchl�ufen
'           sondern durch 8 Timer
'           TimerSpielerAnimation TimerLinksAnimation TimerMitteAnimation TimerRechtsAnimation
'           TimerEinsammlerLinks TimerEinsammlerMitte TimerEinsammlerRechrs TimerEinsammlerSpieler
'           nicht funktioniert hat TimerLorumAnimation da kommt run time error '340' control array element does not exist
'           L�sung: mit Timer geht es auf allen Computern ungef�hr gleich
'----------------------------------------------------------------------------------------------------------------------------
'Version 11
'14.11.2017 Verbesserung
'           Zustand: Bestwerte l�schen geht bisher nur manuell durch Bearbeiten der Datei RoserlBestwerte.log
'           L�sung: Button 'Bestwerte l�schen' in der Form frmBestwerteZeigen l�scht die Datei RoserlBestwerte.log und liest
'                   sie danach neu ein
'14.11.2017 Fehlerkorrektur
'           Zustand: Spielart=Lorum: Bei Klick auf die bewegte Karte beginnt sie ihren Weg erneut oder setzt ihren Weg �ber die sonstige
'                   Strecke hinaus fort oder l�uft sogar r�ckw�rts und geht dann in eine Schleife
'           L�sung: blnLorumL�uft verhindert die Klick-Wiederholung
'18.11.2017 Zustand: Das Beenden �ber das Schlie�kreuz funktioniert nur selten
'           L�sung: Versuch mit Call term_global - bringt nichts
'----------------------------------------------------------------------------------------------------------------------------
'Version 12
'26.11.2017 kosmetische Verbesserung
'           Zustand: der Kartenstapel von Mitte und Spieler ist nicht zentriert
'           L�sung: Wer richtig rechnen kann ist im Vorteil. MitteStapelLeft  und SpielerStapelLeft war falsch ausgerechnet.
'27.11.2017 kosmetische Verbesserung
'           Zustand: Wenn bei Spielart <> Lorum die letzte Karte bewegt wird, kann man solange wie die Bewegung anh�lt, beide Karten
'                   sehen. Die zugedeckte letzte und die bewegte letzte Karte
'           L�sung: Sobald die Bewegung beginnt, wird cdLinks(MyIndex).Visible = False - ebenso cdMitte und cdRechts
'28.11.2017 kosmetische Verbesserung
'           Zustand: Bei Klick auf die Spielerkarte ist die aktuelle Reihenfolge
'                   Spielerkarte bewegen
'                   hinterste Entladen
'                   Spielerkarten NeuAuflegen
'                   Mit anderen Worten, man sieht gleichzeitig die bewegte und die unbewegte liegt noch da
'           L�sung: Reihenfolge vertauschen
'                   hinterste Entladen
'                   Spielerkarten NeuAuflegen
'                   Spielerkarte bewegen
'28.11.2017 Verbesserung der Intelligenz
'           Zustand: Spielart=Ober -> Mitte hat Eichel Unter und Eichel Ober -> nimmt den Stich mit Eichel Ober
'                   -> spielt danach Eichel Unter aus
'           L�sung: 'Function VorbehandlungOber' soll nur nach Ass oder K�nig suchen
'28.11.2017 Verbesserung der Intelligenz
'           'Function VorbehandlungRoterK�nig' soll den K�nig ignorieren bei der Suche nach der h�chsten Karte
'29.11.2017 Fehlerkorrektur
'           Zustand: Bei Spielwiederholung mit Strg+N sind manche Schalter nicht zur�ckgestellt
'           L�sung: blnAlleOberGespielt=False blnRKGespielt=False blnAlleRotenGespielt=False
'08.12.2017 Verbesserung der Intelligenz
'           Zustand: Spielart=RoterK�nig - Rechts ist Ausspieler und spielt RotAcht, Spieler spielt RotAss, links hat den RoterK�nig,
'                   aber spielt statt dessen die RotZehn, obwohl RotAss auf dem Tisch liegt
'           L�sung: Ich brauche noch einen Schalter blnRotAssAufDemTisch
'09.12.2017 kosmetische Verbesserung
'           Zustand: Es ist w�nschenswert am Ende von Spielart=Lorum nach der letzten Msgbox, die mit Enter-Taste beantwortet wird,
'                   auch noch mit der Enter-Taste das Spiel zu beenden, anstelle auf die Form klicken zu m�ssen
'           L�sung: Ein Dr�cken der Enter-Taste wird weitergeleitet zu Form_Click
'10.12.2017 Verbesserung
'           Zustand: Manchmal w�re es zum Fehlerfinden gut, wenn ich in die Kartenverteilung eingreifen k�nnte
'           Anstelle von Mischen teile ich die Karten in fester Reihenfolge aus
'           L�sung: In Prozedur NichtMischen ist eine feste Reihenfolge vorgegeben
'                   gblnNichtMischen = True setzen, wenn Prozedur NichtMischen aufgerufen werden soll
'17.12.2017 Fehlerkorrektur
'           Zustand: Spielart=LorumMitte hat soeben Gr�nOber gespielt, es ist eine MsgBox offen. Bei Dr�cken der Enter-Taste(oder Klick auf OK)
'                   wird immer wieder der Gr�nOber gespielt
'           L�sung: Der Gr�nOber wird immer wieder gespielt, obwohl er .GespieltVon = 9 h�tte sein m�ssen, war es aber nicht
'                   EntladenLorum muss ge�ndert werden; Aufrufparameter gespielteKarte entf�llt
'                   ich suche bis ich die gespielteKarte gefunden habe, die mit LorumAuflegeFarbe und LorumStartWert �bereinstimmt
'                   LorumAssAblegen muss ge�ndert werden
'                   LorumAbw�rtsAblegen muss ge�ndert werden
'19.12.2017 Verbesserung
'           Zustand: G�nter Greif meckert: Wenn schon ein Abbrechen nach 1 oder 2 oder 3 Durchg�ngen angeboten wird, sollte auch ein
'                   Wiedereinsetzen am Abbruchpunkt m�glich sein
'           L�sung: Abbrechen nach 1 oder 2 oder 3 Durchg�ngen wird nicht mehr angeboten. Stattdessen bringe ich die ge�nderte MsgBox
'                   "Die n�chste Spielrunde beginnt.
'                   "Du mu�t zum Weiterspielen auf den roten Buchstaben klicken"
'19.12.2017 Verbesserung
'           Zustand: Das Beenden �ber das Schlie�kreuz funktioniert nur selten
'           L�sung: Schuld ist 'cancel = true' in frmSpielBogen
'                   frmSpielBogen wird �berhaupt nicht mehr entladen(es gibt kein Schlie�kreuz), h�chstens Me.Hide
'19.12.2017 Verbesserung
'           Zustand: Es scheint so als ob h�ufig dieselben Karten verteilt w�rden. Besonders nach Start -> Beenden -> neuer Start
'           L�sung: Anstelle von Randomize Timer -> Randomize
'19.12.2017 Fehlerkorrektur Nachbesserung zu 08.12.2017 '08.12.2017 19.12.2017
'           Zustand: Spielart=RoterK�nig bringt Laufzeitfehler '9' Index au�erhalb...
'           L�sung: falsch ist ...Gespielt.CardType = AKType
'                   richtig ist ...Gespielt.CardType = 2
'18.01.2018 Verbesserung
'           Zustand: Die �nderung der Schriftgr��e wirkt erst nach Schlie�en von frmFontSize
'           L�sung: Sofort nach Klick auf klein/mittel/gro� ist die Schriftgr��en�nderung sichtbar
'
'----------------------------------------------------------------------------------------------------------------------------
'Version 13
'16.02.2018 Verbesserung mit Multi Monitor
'           Zustand: G�nter meckert, dass es schon mal besser ging
'                   frmMsgBox kommt immer auf dem Startmonitor, auch wenn frmRoserl auf den Zweitmonitor verschoben wurde
'                   frmFontSize und frmLetztenStichZeigen und frmBestWerteZeigen kommen korrekt auf dem Zweitmonitor
'                   frmSpielBogen kommt immer auf dem Startmonitor
'           L�sung: frmMsgBox kommt jetzt auf dem Zweitmonitor
'                   frmSpielbogen muss einmal manuell verschoben werden, danach bleibt es auf dem Zweitmonitor
'                   Wenn frmMsgBox auf dem falschen Monitor kommt, meist bei oder nach Lorum, hilft Verschieben auf den anderen Monitor
'           Begleiterscheinung:
'                   MsgBox 'falsch gespielt bei Rot' oder 'Du musst Farbe zugeben' kommt �ber der betreffenden Karte.
'                   Das passiert aber nur in der exe, nicht in der IDE
'           Eventuell weitere Verbesserung im Multimonitor-System:
'                   Wenn Multimonitor-System vorliegt, fragen soll Roserl auf Monitor1 oder Monitor2 gestartet werden
'                   Abgucken bei d:\VISUALBA.SIC\VB6BeispielCode\Dual second Zweit multiple Monitor\NeueForm auf monitor1 oder monitor2 starten\
'19.03.2018 Verbesserung
'           Es gibt ein Men� Spielregeln
'19.03.2018 Verbesserung
'           Das Gerbing-Icon wird durch ein Roserl.ico ersetzt
'----------------------------------------------------------------------------------------------------------------------------
'Version 14
'17.03.2018 Verbesserung G�nter Greif
'           Mit dem Men� 'Beenden' wird Roserl beendet und gleichzeitig wird der Spielstand in die Bin�rdatei roserl.bin gespeichert.
'           Das sind keine Zwischenspielst�nde, sondern der Spielstand nach jeweils einer Roserl-Runde(=6 Spiele).
'           Wenn w�hrend der ersten Runde abgebrochen wird(da gibt es noch kein Runden-Ergebnis), beginnt also Roserl neu
'           Mit roserl.bin ist es m�glich, bei einem abgebrochenen Spielstand weiterzuspielen. Wenn das Programm eine roserl.bin entdeckt,
'           wird der Nutzer gefragt, ob er mit dem Spielstand zu Beginn der abgebrochenen Runde weiterspielen will(mit Rot),
'           oder ob er neu beginnen will. Nach dem Einlesen von roserl.bin wird diese gel�scht.
'           Wer in der 2./3./4. Runde der Rot-Ausspier ist, steht fest.
'----------------------------------------------------------------------------------------------------------------------------
'Version 15
'13.06.2018 Verbesserung G�nter Greif
'           Es wird nicht nur der Spielstand nach jeweils einer Runde eingetragen, sondern auch der Spielstand aller in dieser Runde
'           gespielten Einzelspiele.
'           Es bleibt aber dabei, dass nur Spielst�nde nach vollen Runden gespeichert werden
'           Das Gerbing-Icon wird in s�mtlichen Formularen durch ein Roserl48Sieben.ico ersetzt
'10.10.2018 Verbesserung
'           Zustand: Im Windows 10 kommt Roserl verlangt eine Bildschirmaufl�sung von mindestens 1920x1060
'           L�sung: Ich erg�nze:gefunden sreenwidth=xxxx screenheight=yyyy
'                   Ich erhalte:
'                   Roserl verlangt eine Bildschirmaufl�sung von mindestens 1920x1060
'                   gefunden sreenwidth=1536 screenheight=960
'           das liegt daran, dass ich Bildschirm Skalierung und Anordnung auf 125% eingestellt habe.
'           Gegenma�nahme: Eigenschaften der exe �ffnen -> Kompatibilit�t -> Hohe DPI-Einstellungen �ndern -> Hohe DPI Skalierung �berschreiben
'                   -> Skalierung wird durchgef�hrt von -> Anwendung
'----------------------------------------------------------------------------------------------------------------------------
'20.03.2020 Verbesserung
'           Zustand: Es gibt Beschwerden, dass angeblich zu oft die gleichen Karten verteilt werden(G�nter Greif)
'           L�sung: Es wird nur einmal zu Programmbeginn Randomize gemacht
'----------------------------------------------------------------------------------------------------------------------------
'Version 15.0.1
'26.07.2020 Verbesserung
'           Zustand: Manchmal �rgert mich meine bl�de Spielweise, wenn ich ein Spiel schlecht beendet habe. Dann w�rde ich gern
'                   die Chance erhalten das selbe Spiel noch einmal zu spielen, ohne dass mein schlechtes Spiel angerechnet wird.
'           L�sung: Ich muss die bisherige L�sung Strg+N anpassen
'                   Ich bringe bei Spiel-Ende keine MsgBox mit dem Spielstand sondern die neue Form
'                   frmSpielEnde. Dort gibt es den Button 'Spiel wiederholen' oder 'Spiel fortsetzen'
'----------------------------------------------------------------------------------------------------------------------------
'Version 15.0.1
'28.07.2020 kosmetische Verbesserung
'           Zustand: Bei Eingabe von 'maloi' decken die Mitspieler ihre Karten auf. Bisher sind diese unsortiert und somit un�bersichtlich.
'                   Ich will sortierte Darstellung.
'           L�sung: Die Prozeduren LinksKartenNeu, MitteKartenNeu, RechtsKartenNeu werden neu geschrieben
'           --aber--
'           bei Lorum wird bei Links und auch bei Rechts bei der gemovten Karte ein falsches Kartenbild angezeigt(bei Mitte ist es richtig)
'           das geht nur weg, wenn ich in Prozedur MoveKarteLorum
'           Facemode = NachOben �ndere in FaceMode = R�ckseite
'----------------------------------------------------------------------------------------------------------------------------
'Version 15.0.1
'22.10.2020 Nachbesserung zum 28.07.2020
'           Zustand: Die Mitte-Karten werden nach dem Ausspielen nicht in der Bildschirmmitte ausgerichtet, so wie beim Spieler
'           L�sung: Programmierfehler es wurde mit cdSpieler.Count gerechnet anstelle von
'                   cdMitte.Count
'----------------------------------------------------------------------------------------------------------------------------
'Version 15.0.2
'19.05.2021 Nachbesserung zum 26.07.2020
'           Zustand: Seit ich ein schlechtes Spiel wiederholen kann, gewinne ich fast alle Spiele. Das macht eigentlich keinen Spa�.
'                   Ich will eine Variante wo der Button 'Spiel wiederholen' sichtbar ist und eine wo der Button nicht sichtbar ist.
'           L�sung:     #If Spielwiederholen = 0 Then                   '0=nein
'                           btnSpielWiederholen.Visible = False
'                       #Else
'                           btnSpielWiederholen.Visible = True
'                       #End If
'                   Projekt -> Eigenschaften von Roserl -> Erstellen -> Argumente f�r bedingte Kompilierung -> Spielwiederholen = 0
'                       hei�t ohne den Button 'Spiel wiederholen' erzeugt die RoserlO.exe
'                   Projekt -> Eigenschaften von Roserl -> Erstellen -> Argumente f�r bedingte Kompilierung -> Spielwiederholen = 1
'                       hei�t mit Button 'Spiel wiederholen' erzeugt die RoserlM.exe
'                   -----------------------------------------------------------------------
'                       #If Spielwiederholen = 0 Then                   '0=nein                                 'Gerbing 19.05.2021
'                           Me.Caption = "RoserlO"
'                       #Else
'                           Me.Caption = "RoserlM"
'                       #End If
'
'
'
'
'
'
'
'
'
'
'
'Probleme
'20.09.2017 Beim ersten frmRoserl_Resize r�ckt die Form zu weit nach oben. Ich finde keine Gegenma�nahme �bers Programm, aber es hilft
'           einfach am Rand zu ziehen.
'20.09.2017 bei G�nter Greif: Roserl wird beendet und beim Neustart kommt:
'           MsgBox "Roserl verlangt eine Bildschirmaufl�sung von mindestens 1920x1060." ist ganz oben links in der Ecke
'           und von frmRoserl ist nichts zu sehen. Das Hintergrundbild ist zu sehen.
'           G�nter benutzt ein MultiMonitor-System.
'           User an- und abmelden hilft. Dann geht's normal weiter.
'           Ich kann es nicht reproduzieren.
'19.01.2018 Die Bestwerte werden erst eingetragen, nachdem nach der 4.Runde weiterspielen gew�hlt wurde und nachdem der erste Stich Rot
'           ausgespielt wurde.
'           Ich kann es nicht reproduzieren.
'17.03.2018 Roserl.pdb wird erzeugt durch Compilieren mit H�kchen bei 'Debug-Informationen f�r symbolischen Debugger generieren'
'           Bei ungekl�rten Laufzeitfehlern, die nur in der exe passieren, kann man dann mit dem C++ Debugger weitere Informationen bekommen
    
    Public LinksPunkte As Integer
    Public RechtsPunkte As Integer
    Public MittePunkte As Integer
    Public SpielerPunkte As Integer
    
    Private Type Rang
        RangPosition As Integer
        Vergeben As Boolean
    End Type

    Dim Rangliste() As Rang
    
    Const MIN_DIMEN = 6500
    
    Dim StartpunktXLinks As Long                                            'Gerbing 20.09.2017
    Dim StartPunktYLinks As Long
    Dim StartpunktXMitte As Long
    Dim StartPunktYMitte As Long
    Dim StartpunktXRechts As Long
    Dim StartPunktYRechts As Long
    Dim StartpunktXSpieler As Long
    Dim StartPunktYSpieler As Long
    Dim EndpunktXLinks As Long                                              'Gerbing 20.09.2017
    Dim EndPunktYLinks As Long
    Dim EndpunktXMitte As Long
    Dim EndPunktYMitte As Long
    Dim EndpunktXRechts As Long
    Dim EndPunktYRechts As Long
    Dim EndpunktXSpieler As Long
    Dim EndPunktYSpieler As Long
    
    Dim StartPunktXLorum As Long                                            'Gerbing 25.09.2017
    Dim StartPunktYLorum As Long
    Dim ZielPunktLeftLorum As Long
    Dim ZielPunktTopLorum As Long
    Dim WegX As Long
    Dim WegY As Long
    Dim AbsWegX As Long
    Dim AbsWegY As Long
    Dim strAbleger As String
    Dim AssPosition As Integer

    Dim gespielteKarteIndex As Integer
    
    Dim blnSpielerWarDran As Boolean
    Dim blnLinksWarDran As Boolean
    Dim blnMitteWarDran As Boolean
    Dim blnRechtsWarDran As Boolean
    
    Dim LinksStapelTop As Integer
    Dim MitteStapelTop As Integer
    Dim RechtsStapelTop As Integer
    Dim SpielerStapelTop As Integer
    
    Dim RechtsStapelLeft As Integer
    Dim LinksStapelLeft As Integer
    Dim MitteStapelLeft As Integer
    Dim SpielerStapelLeft As Integer
    
    Dim blnMoveL�uft As Boolean
    Dim blnRKGespielt As Boolean                    'RoterK�nigGespielt
    Dim blnAlleOberGespielt As Boolean              'Gerbing 12.09.2017
    Dim blnAlleRotenGespielt As Boolean             'Gerbing 12.09.2017
    Dim blnSchellZEM As Boolean                     'Schell Zum Ersten Mal
    Dim blnRotZEM As Boolean                        'Rot Zum Ersten Mal
    Dim blnGr�nZEM As Boolean                       'Gr�n Zum Ersten Mal
    Dim blnEichelZEM As Boolean                     'Eichel Zum Ersten Mal
    Dim blnStapelBBenutzen As Boolean
    Dim Aufrufz�hlerSpieler As Integer                                  'Gerbing 09.11.2017
    Dim Aufrufz�hlerLinks As Integer                                    'Gerbing 09.11.2017
    Dim Aufrufz�hlerMitte As Integer                                    'Gerbing 09.11.2017
    Dim Aufrufz�hlerRechts As Integer                                   'Gerbing 09.11.2017
    Dim Aufrufz�hlerEinsammler As Integer                               'Gerbing 09.11.2017
    Dim Aufrufz�hlerLorum As Integer                                    'Gerbing 09.11.2017
    Dim LorumPosition As Integer                                        'Gerbing 09.11.2017
    Dim blnloruml�uft As Boolean                                        'Gerbing 14.11.2017
    
    Dim LorumStartWert As Integer
    Dim LorumStartPosition As Integer
    Dim LorumAuflegeFarbe As Integer
    Dim LorumAufgelegtZ�hler As Integer
    Dim blnLorumSpielende As Boolean
    Dim blnErsteKommtVonSpieler As Boolean
    Dim blnErsteKommtVonLinks As Boolean
    Dim blnErsteKommtVonMitte As Boolean
    Dim blnErsteKommtVonRechts As Boolean
    Public PlayerTurnRoserl As Players
    Dim LorumGewinner As Players
    Dim nDoEvents As Integer
    Dim iMaxDoEvents As Integer
    Dim iMaxMove As Integer
    Dim iMulti As Integer
    Dim XMulti As Double
    Dim YMulti As Double
    
    Private Declare Sub GetSystemTime Lib _
        "kernel32" (lpSystemTime As SYSTEMTIME)
    Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
    End Type

Private Sub MindestAufl�sung()
'    screenWidth = GetDeviceCaps(Me.hdc, HORZRES)
'    screenHeight = GetDeviceCaps(Me.hdc, VERTRES)
    screenWidth = GetSystemMetrics(SM_CXSCREEN)         'Gerbing 10.10.2018
    screenHeight = GetSystemMetrics(SM_CYSCREEN)        'Gerbing 10.10.2018


    If screenWidth < 1920 Or screenHeight < 1060 Then
        gblnMsgBoxAntwort = True
        gstrMsg = "Roserl verlangt eine Bildschirmaufl�sung von mindestens 1920x1060." & vbNewLine
        gstrMsg = gstrMsg & "gefunden screenWidth=" & screenWidth & " screenHeight=" & screenHeight & vbNewLine 'Gerbing 10.10.2018
        gstrMsg = gstrMsg & "Willst Du trotzdem spielen?"
        frmMsgBox.Show 1
        If glngAntwort = vbNo Then
            End
        End If
    End If
End Sub

Private Sub WelchenAblegestapelBenutzen(LorumAuflegeFarbe As Integer)
    blnStapelBBenutzen = False
    'wenn bereits 4 Karten im StapelA abgelegt sind
    If cdRoserlSchellA.Count >= 5 And LorumAuflegeFarbe = 1 Then blnStapelBBenutzen = True
    If cdRoserlRotA.Count >= 5 And LorumAuflegeFarbe = 2 Then blnStapelBBenutzen = True
    If cdRoserlGr�nA.Count >= 5 And LorumAuflegeFarbe = 3 Then blnStapelBBenutzen = True
    If cdRoserlEichelA.Count >= 5 And LorumAuflegeFarbe = 4 Then blnStapelBBenutzen = True
    'wenn die h�chste Karte ein Ass ist
    If cdRoserlSchellA(0).Visible = True Then
        If cdRoserlSchellA(cdRoserlSchellA.Count - 2).CardValue = 8 And LorumAuflegeFarbe = 1 Then blnStapelBBenutzen = True
    End If
    If cdRoserlRotA(0).Visible = True Then
        If cdRoserlRotA(cdRoserlRotA.Count - 2).CardValue = 8 And LorumAuflegeFarbe = 2 Then blnStapelBBenutzen = True
    End If
    If cdRoserlGr�nA(0).Visible = True Then
        If cdRoserlGr�nA(cdRoserlGr�nA.Count - 2).CardValue = 8 And LorumAuflegeFarbe = 3 Then blnStapelBBenutzen = True
    End If
    If cdRoserlEichelA(0).Visible = True Then
        If cdRoserlEichelA(cdRoserlEichelA.Count - 2).CardValue = 8 And LorumAuflegeFarbe = 4 Then blnStapelBBenutzen = True
    End If
    'wenn zB die h�chste Karte von SchellA und die niedrigste Karte von SchellAss aufeinander folgen
    If cdRoserlSchellAss(0).Visible = True And cdRoserlSchellA(0).Visible = True Then
        If cdRoserlSchellA(cdRoserlSchellA.Count - 2).CardValue = cdRoserlSchellAss(cdRoserlSchellAss.Count - 1).CardValue - 1 And LorumAuflegeFarbe = 1 Then blnStapelBBenutzen = True
    End If
    If cdRoserlRotAss(0).Visible = True And cdRoserlRotA(0).Visible = True Then
        If cdRoserlRotA(cdRoserlRotA.Count - 2).CardValue = cdRoserlRotAss(cdRoserlRotAss.Count - 1).CardValue - 1 And LorumAuflegeFarbe = 2 Then blnStapelBBenutzen = True
    End If
    If cdRoserlGr�nAss(0).Visible = True And cdRoserlGr�nA(0).Visible = True Then
        If cdRoserlGr�nA(cdRoserlGr�nA.Count - 2).CardValue = cdRoserlGr�nAss(cdRoserlGr�nAss.Count - 1).CardValue - 1 And LorumAuflegeFarbe = 3 Then blnStapelBBenutzen = True
    End If
    If cdRoserlEichelAss(0).Visible = True And cdRoserlEichelA(0).Visible = True Then
        If cdRoserlEichelA(cdRoserlEichelA.Count - 2).CardValue = cdRoserlEichelAss(cdRoserlEichelAss.Count - 1).CardValue - 1 And LorumAuflegeFarbe = 4 Then blnStapelBBenutzen = True
    End If
End Sub

Private Sub cdLinks_Click(Index As Integer)
    Dim rc As Integer
    
    If Spielart = Lorum Then
        If blnloruml�uft Then Exit Sub                                      'Gerbing 14.11.2017
        If PlayerTurnRoserl = Links Then
            blnloruml�uft = True                                            'Gerbing 14.11.2017
            PlayerTurn = Links
            'Untersuche an welchen Schell Rot Gr�n Eichel Stapel an oder abgelegt werden soll
            LorumAuflegeFarbe = cdLinks(Index).CardType
            Call WelchenAblegestapelBenutzen(LorumAuflegeFarbe)
            rc = SucheBesteKarteLorumAusspieler
            blnloruml�uft = False                                           'Gerbing 14.11.2017
        End If
    End If
End Sub

Private Sub cdMitte_Click(Index As Integer)
    Dim rc As Integer
    
    If Spielart = Lorum Then
        If blnloruml�uft Then Exit Sub                                      'Gerbing 14.11.2017
        If PlayerTurnRoserl = Mitte Then
            blnloruml�uft = True                                            'Gerbing 14.11.2017
            PlayerTurn = Mitte
            'Untersuche an welchen Schell Rot Gr�n Eichel Stapel an oder abgelegt werden soll
            LorumAuflegeFarbe = cdMitte(Index).CardType
            Call WelchenAblegestapelBenutzen(LorumAuflegeFarbe)
            rc = SucheBesteKarteLorumAusspieler
            blnloruml�uft = False                                           'Gerbing 14.11.2017
        End If
    End If
End Sub

Private Sub cdRechts_Click(Index As Integer)
    Dim rc As Integer
    
    If Spielart = Lorum Then
        If blnloruml�uft Then Exit Sub                                      'Gerbing 14.11.2017
        If PlayerTurnRoserl = Rechts Then
            blnloruml�uft = True                                            'Gerbing 14.11.2017
            PlayerTurn = Rechts
            'Untersuche an welchen Schell Rot Gr�n Eichel Stapel an oder abgelegt werden soll
            LorumAuflegeFarbe = cdRechts(Index).CardType
            Call WelchenAblegestapelBenutzen(LorumAuflegeFarbe)
            rc = SucheBesteKarteLorumAusspieler
            blnloruml�uft = False                                           'Gerbing 14.11.2017
        End If
    End If
End Sub

Private Sub cdSpieler_Click(Index As Integer)
    Dim I As Integer
    Dim rc As Integer
    Dim GKType As Integer
    Dim GKValue As Integer

    If Spielart = Lorum Then
        If blnloruml�uft Then Exit Sub                                      'Gerbing 14.11.2017
        If PlayerTurnRoserl = Spieler Then
            blnloruml�uft = True                                            'Gerbing 14.11.2017
            PlayerTurn = Spieler
            'Untersuche an welchen Schell Rot Gr�n Eichel Stapel diese Karte an oder abgelegt werden kann, StapelA oder StapelB
            LorumAuflegeFarbe = cdSpieler(Index).CardType
            Call WelchenAblegestapelBenutzen(LorumAuflegeFarbe)
            LorumStartWert = cdSpieler(Index).CardValue
            LorumStartPosition = Index + 1
            LorumAufgelegtZ�hler = 1
            'wenn der Spieler ein Ass ausspielt
            If cdSpieler(Index).CardValue = 8 Then                                     '8=Ass
                rc = LorumAssAblegen()
                blnloruml�uft = False                                           'Gerbing 14.11.2017
                Exit Sub
            End If
            '-----------------------
            LorumAuflegeFarbe = cdSpieler(Index).CardType
            LorumStartWert = cdSpieler(Index).CardValue
            LorumStartPosition = Index + 1
            LorumAufgelegtZ�hler = 1
            'wenn der Spieler eine m�gliche Abw�rts-Anlegekarte ausspielt
            If cdSpieler(Index).CardValue < 8 Then                                     '<8=unter dem Ass
                rc = LorumAbw�rtsAblegen()
                If rc <> 0 Then
                    'Abw�rtskarte gefunden
                    blnloruml�uft = False                                       'Gerbing 14.11.2017
                    Exit Sub
                Else
                    'Abw�rtskarte nicht gefunden
'                    gblnMsgBoxAntwort = True
'                    gstrMsg = "Diese Karte ist keine Abw�rts-Karte, willst Du sie einzeln ablegen?"
'                    frmMsgBox.Show 1
'                    If glngAntwort = vbYes Then
                        blnStapelBBenutzen = False
                        LorumAuflegeFarbe = cdSpieler(Index).CardType
                        LorumStartWert = cdSpieler(Index).CardValue
                        LorumStartPosition = Index + 1
                        LorumAufgelegtZ�hler = 1
                        blnErsteKommtVonSpieler = True
                        Call LorumEineKarteHinlegen(Index)
                        blnloruml�uft = False                                   'Gerbing 14.11.2017
                        Exit Sub
'                    Else
'                        Exit Sub
'                    End If
                End If
            End If
            blnloruml�uft = False                                               'Gerbing 14.11.2017
        End If
        Exit Sub
    End If
    '-----------------------------------------------------------
    'hier ist Spielart <> Lorum
    If Ausspieler <> Spieler And blnRechtsWarDran = False Then Exit Sub         'Gerbing 08.09.2017
    If blnEingesammelt = False Then Exit Sub                                    'Gerbing 13.08.2017
    If blnNeuesSpielBeginnt = True Then Exit Sub                                'Der Spieler muss auf lblRoserl klicken
    PlayerTurn = Spieler
    gespielteKarteIndex = Index + 1                                             'Gerbing 20.05.2017
    SchlafTimer.Enabled = False
    If IsPlayable(Index) = True Then
        'Wenn die angeklickte Karte gespielt werden darf
        blnEingesammelt = False                                                 'Gerbing 08.09.2017
        GKType = SpielerCards(gespielteKarteIndex).cType
        GKValue = SpielerCards(gespielteKarteIndex).cValue
        Call HintersteEntladen(PlayerTurn, gespielteKarteIndex)
        If cdSpieler.Count <> 1 Then
            Call SpielerKartenNeu                                               'Gerbing 28.11.2017 'Gerbing 16.08.2017
            SpielerGespielt.ZOrder
        End If
        Call EineKarteAblegen(GKType, GKValue)                      'Bewegungsanimation und ausgespielte Karte merken 'Gerbing 28.11.2017
        blnSpielerWarDran = True
        '------------------------------------------------------------
        If Ausspieler = Links Then
            'alle haben schon ausgespielt d.h.eine Karte hingelegt
        End If
        If Ausspieler = Mitte Then
            'Links muss noch ausspielen
            Call LinksPlay
        End If
        If Ausspieler = Rechts Then
            'Links und Mitte m�ssen noch ausspielen
            Call LinksPlay
            Call MittePlay
        End If
        If Ausspieler = Spieler Then
            'Links und Mitte und Rechts m�ssen noch ausspielen
            Call LinksPlay
            Call MittePlay
            Call RechtsPlay
        End If
    End If
End Sub

Private Sub Form_Click()
    'Alle Karten werden eingesammelt in AusgespielteKarten(1 To 32)
    'Prozedur Einsammeln speichert wer den Stich eingenommen hat und ermittelt wer als N�chster ausspielen muss,
    '   das k�nnen sein LinksPlay MittePlay RechtsPlay oder der Spieler
    'Prozedur LinksPlay MittePlay oder RechtsPlay bestimmt, welche Karte als n�chste ausgespielt werden kann oder muss
    'Wenn der Spieler ausspielen muss, warten dass dieser auf eine Karte klickt
    Dim I As Integer
    
    On Error Resume Next
    On Error GoTo 0
    If blnMoveL�uft = True Then Exit Sub
    
    SchlafTimer.Enabled = False                                     'Gerbing 26.09.2017
    If AKIndex = 32 Then
        For I = 0 To cdSpieler.Count - 1                            'Manchmal bleibt beim Spieler die letzte Karte sichtbar
            cdSpieler(I).Visible = False
        Next I
        Call Einsammeln
        Call Spielende
        blnAlleOberGespielt = False                                 'Gerbing 12.09.2017
        Z�hlerOber = 0
        blnAlleRotenGespielt = False
        Z�hlerRote = 0
        Exit Sub
    End If

    If blnLorumSpielende = True Then
        Call Spielende
        Exit Sub
    End If
       
    If blnSpielerWarDran = True And blnLinksWarDran = True And blnMitteWarDran = True And blnLinksWarDran = True Then
        If AKIndex > 3 Then
            Call Einsammeln
        Else
            Exit Sub
        End If
        'das passiert nach dem Einsammeln
        If blnRKGespielt = True Then
            blnRKGespielt = False
            Call Spielende
            'MsgBox "Der rote K�nig ist gespielt, damit endet das Spiel"
            Exit Sub
        End If
        If blnAlleOberGespielt = True Then                          'Gerbing 12.09.2017
            blnAlleOberGespielt = False
            Z�hlerOber = 0
            Call Spielende
            'MsgBox "Alle Ober sind gespielt, damit endet das Spiel"
            Exit Sub
        End If
        If blnAlleRotenGespielt = True Then                         'Gerbing 12.09.2017
            blnAlleRotenGespielt = False
            Z�hlerRote = 0
            Call Spielende
            'MsgBox "Alle Roten sind gespielt, damit endet das Spiel"
            Exit Sub
        End If
        If Ausspieler = Links Then
            Call LinksPlay
            Call MittePlay
            Call RechtsPlay
            'warten auf cdSpieler_Click
        End If
        If Ausspieler = Mitte Then
            Call MittePlay
            Call RechtsPlay
            'warten auf cdSpieler_Click
        End If
        If Ausspieler = Rechts Then
            Call RechtsPlay
            'warten auf cdSpieler_Click
        End If
        If Ausspieler = Spieler Then
            'warten auf cdSpieler_Click
        End If
    End If
End Sub

Private Sub EntladenLorum(PlayerTurn As Players)
    'EntladenLorum kommt nur bei Spielart = Lorum dran
    'Ich durchsuche alle Spieler Links Mitte Rechts Cards(mit Ausnahme der mit Kennzeichen .GespieltVon = 9) bis
    'ich die gespielteKarte gefunden habe, die mit LorumAuflegeFarbe und LorumStartWert �bereinstimmt
    'z�hle die mit Kennzeichen .GespieltVon=9
    'entlade sie und setze das Kennzeichen .GespieltVon=9
    
    Dim I As Integer
    Dim n As Integer
    Dim strSpieler As String
    
    n = 0
    If PlayerTurn = Spieler Then
        For I = 1 To 8
            If SpielerCards(I).cGespieltVon <> 9 Then
                If SpielerCards(I).cType = LorumAuflegeFarbe And SpielerCards(I).cValue = LorumStartWert Then   'Gerbing 17.12.2017
                'If I = gespielteKarte Then
                    SpielerCards(I).cGespieltVon = 9
                    If I <> 1 Then                          'die erste Karte kann nicht entladen werden
                        Unload cdSpieler(I - 1)
                    Else
                        cdSpieler(I - 1).Visible = False
                    End If
                    Call SpielerKartenNeu
                End If
            Else
                n = n + 1                                   'Anzahl Karten mit Kennzeichen .GespieltVon=9
            End If
        Next I
        End If
    If PlayerTurn = Links Then
        For I = 1 To 8
            If LinksCards(I).cGespieltVon <> 9 Then
                If LinksCards(I).cType = LorumAuflegeFarbe And LinksCards(I).cValue = LorumStartWert Then   'Gerbing 17.12.2017
                'If I = gespielteKarte Then
                    LinksCards(I).cGespieltVon = 9
                    If I <> 1 Then                          'die erste Karte kann nicht entladen werden
                        Unload cdLinks(I - 1)
                    Else
                        cdLinks(I - 1).Visible = False
                    End If
                    Call LinksKartenNeu
                End If
            Else
                n = n + 1                                   'Anzahl Karten mit Kennzeichen .GespieltVon=9
            End If
        Next I
    End If
    If PlayerTurn = Mitte Then
        For I = 1 To 8
            If MitteCards(I).cGespieltVon <> 9 Then
                If MitteCards(I).cType = LorumAuflegeFarbe And MitteCards(I).cValue = LorumStartWert Then   'Gerbing 17.12.2017
                'If I = gespielteKarte Then
                    MitteCards(I).cGespieltVon = 9
                    If I <> 1 Then                          'die erste Karte kann nicht entladen werden
                        Unload cdMitte(I - 1)
                    Else
                        cdMitte(I - 1).Visible = False
                    End If
                    Call MitteKartenNeu
                End If
            Else
                n = n + 1                                   'Anzahl Karten mit Kennzeichen .GespieltVon=9
            End If
        Next I
    End If
    If PlayerTurn = Rechts Then
        For I = 1 To 8
            If RechtsCards(I).cGespieltVon <> 9 Then
                If RechtsCards(I).cType = LorumAuflegeFarbe And RechtsCards(I).cValue = LorumStartWert Then 'Gerbing 17.12.2017
                'If I = gespielteKarte Then
                    RechtsCards(I).cGespieltVon = 9
                    If I <> 1 Then                          'die erste Karte kann nicht entladen werden
                        Unload cdRechts(I - 1)
                    Else
                        cdRechts(I - 1).Visible = False
                    End If
                    Call RechtsKartenNeu
                End If
            Else
                n = n + 1                                   'Anzahl Karten mit Kennzeichen .GespieltVon=9
            End If
        Next I
    End If
    '------------------------------------------------------------
    'hier ist Spielart = Lorum zu Ende
    If PlayerTurn = Links And n >= 7 Then
        strSpieler = "Links"
        blnLorumSpielende = True
        LorumGewinner = Links
        On Error Resume Next
        For I = 0 To 7
            cdLinks(I).Visible = False                      'Gerbing 15.08.2017
        Next
        On Error GoTo 0
    ElseIf PlayerTurn = Mitte And n >= 7 Then
        strSpieler = "Mitte"
        blnLorumSpielende = True
        LorumGewinner = Mitte
        On Error Resume Next
        For I = 0 To 7
            cdMitte(I).Visible = False                      'Gerbing 15.08.2017
        Next
        On Error GoTo 0
    ElseIf PlayerTurn = Rechts And n >= 7 Then
        strSpieler = "Rechts"
        blnLorumSpielende = True
        LorumGewinner = Rechts
        On Error Resume Next
        For I = 0 To 7
            cdRechts(I).Visible = False                      'Gerbing 15.08.2017
        Next
        On Error GoTo 0
    ElseIf PlayerTurn = Spieler And n >= 7 Then
        strSpieler = "Spieler"
        blnLorumSpielende = True
        LorumGewinner = Spieler
        On Error Resume Next
        For I = 0 To 7
            cdSpieler(I).Visible = False                      'Gerbing 15.08.2017
        Next
        On Error GoTo 0
    End If
    If blnLorumSpielende = True And Spielart = Lorum Then
        'das Spiel ist zu Ende
        gstrMsg = "Lorum wurde von " & strSpieler & vbNewLine
        gstrMsg = gstrMsg & "gewonnen"
        frmMsgBox.Show 1
        Exit Sub
    End If
End Sub

Public Sub HintersteEntladen(PlayerTurn As Players, gespielteKarte As Integer)
    'HintersteEntladen kommt nur bei Spielart <> Lorum dran
    
    Dim I As Integer
    Dim strSpieler As String
    
    'es werden alle CardType und CardValue ab gespielteKarte um eine Karte nach vorn verschoben  damit die hinterste entladen werden kann
    If PlayerTurn = Links Then
        For I = gespielteKarte To UBound(LinksCards) - 1
            LinksCards(I) = LinksCards(I + 1)
            cdLinks(I - 1).SetCardValue cdLinks(I).CardType, cdLinks(I).CardValue 'Gerbing 30.06.2017
        Next I
    End If
    If PlayerTurn = Mitte Then
        For I = gespielteKarte To UBound(MitteCards) - 1
            MitteCards(I) = MitteCards(I + 1)
            cdMitte(I - 1).SetCardValue cdMitte(I).CardType, cdMitte(I).CardValue 'Gerbing 30.06.2017
        Next I
    End If
    If PlayerTurn = Rechts Then
        For I = gespielteKarte To UBound(RechtsCards) - 1
            RechtsCards(I) = RechtsCards(I + 1)
            cdRechts(I - 1).SetCardValue cdRechts(I).CardType, cdRechts(I).CardValue 'Gerbing 30.06.2017
        Next I
    End If
    If PlayerTurn = Spieler Then
        For I = gespielteKarte To UBound(SpielerCards) - 1
            SpielerCards(I) = SpielerCards(I + 1)
            cdSpieler(I - 1).SetCardValue cdSpieler(I).CardType, cdSpieler(I).CardValue 'Gerbing 30.06.2017
        Next I
    End If
    '-------------------------------------------------
    If PlayerTurn = Links Then
        If cdLinks.Count = 1 Then
            cdLinks(cdLinks.Count - 1).Visible = False
        Else
            cdLinks(cdLinks.Count - 1).Visible = False
            Unload cdLinks(cdLinks.Count - 1)                               'hinterste entladen
        End If
    End If
    If PlayerTurn = Mitte Then
        If cdMitte.Count = 1 Then
            cdMitte(cdMitte.Count - 1).Visible = False
        Else
            cdMitte(cdMitte.Count - 1).Visible = False
            Unload cdMitte(cdMitte.Count - 1)                               'hinterste entladen
        End If
    End If
    If PlayerTurn = Rechts Then
        If cdRechts.Count = 1 Then
            cdRechts(cdRechts.Count - 1).Visible = False
        Else
            cdRechts(cdRechts.Count - 1).Visible = False
            Unload cdRechts(cdRechts.Count - 1)                             'hinterste entladen
        End If
    End If
    If PlayerTurn = Spieler Then
        If cdSpieler.Count = 1 Then
            cdSpieler(cdSpieler.Count - 1).Visible = False
        Else
            cdSpieler(cdSpieler.Count - 1).Visible = False
            Unload cdSpieler(cdSpieler.Count - 1)                           'hinterste entladen
        End If
    End If
    '------------------------------------------------------------
    If PlayerTurn = Links And UBound(LinksCards) = 1 Then Exit Sub
    If PlayerTurn = Mitte And UBound(MitteCards) = 1 Then Exit Sub
    If PlayerTurn = Rechts And UBound(RechtsCards) = 1 Then Exit Sub
    If PlayerTurn = Spieler And UBound(SpielerCards) = 1 Then Exit Sub
    '------------------------------------------------------------
    'das Spiel ist nicht zu Ende
    If PlayerTurn = Links Then
        ReDim Preserve LinksCards(1 To UBound(LinksCards) - 1)
        'Call LinksKartenNeuAuflegen
        Call LinksKartenNeu
    ElseIf PlayerTurn = Mitte Then
        ReDim Preserve MitteCards(1 To UBound(MitteCards) - 1)
        Call MitteKartenNeu                                     'Gerbing 23.06.2017
    ElseIf PlayerTurn = Rechts Then
        ReDim Preserve RechtsCards(1 To UBound(RechtsCards) - 1)
        Call RechtsKartenNeu                                    'Gerbing 23.06.2017
    ElseIf PlayerTurn = Spieler Then
        ReDim Preserve SpielerCards(1 To UBound(SpielerCards) - 1)
        Call SpielerKartenNeu                                    'Gerbing 23.06.2017
        SpielerGespielt.ZOrder                                  'Gerbing 20.09.2017
    End If
End Sub

Sub LinksPlay()
    'Entscheiden welche Karte gespielt werden soll oder darf in Prozedur SucheBesteKarte
    '(gespielteKarteIndex) und diese ausspielen
    
    Dim GKType As Integer
    Dim GKValue As Integer
    
    PlayerTurn = Links
    gespielteKarteIndex = SucheBesteKarte                                   'Entscheiden welche Karte gespielt werden soll oder darf
    If Spielart = Lorum Then
        Exit Sub
    End If
    
    SchlafTimer.Enabled = False                                             'Gerbing 26.09.2017
    GKType = LinksCards(gespielteKarteIndex).cType
    GKValue = LinksCards(gespielteKarteIndex).cValue
    Call EineKarteAblegen(GKType, GKValue)                                  'Bewegungsanimation und ausgespielte Karte merken
    Call HintersteEntladen(PlayerTurn, gespielteKarteIndex)
    blnLinksWarDran = True
End Sub

Sub MittePlay()
    'Entscheiden welche Karte gespielt werden soll oder darf in Prozedur SucheBesteKarte
    '(gespielteKarteIndex) und diese ausspielen
    
    Dim GKType As Integer
    Dim GKValue As Integer

    PlayerTurn = Mitte
    gespielteKarteIndex = SucheBesteKarte                                   'Entscheiden welche Karte gespielt werden soll oder darf
    If Spielart = Lorum Then
        Exit Sub
    End If
    
    SchlafTimer.Enabled = False                                             'Gerbing 26.09.2017
    GKType = MitteCards(gespielteKarteIndex).cType
    GKValue = MitteCards(gespielteKarteIndex).cValue
    Call EineKarteAblegen(GKType, GKValue)                                  'Bewegungsanimation und ausgespielte Karte merken
    Call HintersteEntladen(PlayerTurn, gespielteKarteIndex)
    blnMitteWarDran = True
End Sub

Sub RechtsPlay()
    'Entscheiden welche Karte gespielt werden soll oder darf in Prozedur SucheBesteKarte
    '(gespielteKarteIndex) und diese ausspielen
    
    Dim GKType As Integer
    Dim GKValue As Integer

    PlayerTurn = Rechts
    gespielteKarteIndex = SucheBesteKarte                                   'Entscheiden welche Karte gespielt werden soll oder darf
    If Spielart = Lorum Then
        Exit Sub
    End If
    
    GKType = RechtsCards(gespielteKarteIndex).cType
    GKValue = RechtsCards(gespielteKarteIndex).cValue
    Call EineKarteAblegen(GKType, GKValue)                                  'Bewegungsanimation und ausgespielte Karte merken
    Call HintersteEntladen(PlayerTurn, gespielteKarteIndex)
    blnRechtsWarDran = True
    SchlafTimer.Enabled = True                                              'Gerbing 20.09.2017
End Sub

Private Sub EineKarteAblegen(GKType As Integer, GKValue As Integer)
    'Prozedur EineKarteAblegen macht die Bewegungsanimation und merkt sich die ausgespielte Karte f�r die Stichauswertung(Einsammeln)

    Dim I As Integer
    Dim MyIndex As Integer
    
    AKIndex = AKIndex + 1
    MyIndex = gespielteKarteIndex - 1
    '----------------------------------------------------------------
    If PlayerTurn = Spieler Then
        AusgespielteKarten(AKIndex).cType = GKType
        AusgespielteKarten(AKIndex).cValue = GKValue
        AusgespielteKarten(AKIndex).cGespieltVon = PlayerTurn
        StartpunktXSpieler = (ScaleWidth / 2) - NormalWidth / 2
        StartPunktYSpieler = cdSpieler(0).Top                                             'egal welcher index alle sind top gleich
        SpielerGespielt.CardType = GKType
        SpielerGespielt.Width = NormalWidth
        SpielerGespielt.Height = NormalHeight
        SpielerGespielt.CardValue = GKValue
        SpielerGespielt.FaceMode = NachOben
        SpielerGespielt.ZOrder                                                                  'Gerbing 20.09.2017
        blnMoveL�uft = True
        StartMillisek = timeGetTime
        Aufrufz�hlerSpieler = 0
        TimerSpielerAnimation.Enabled = True
        DoEvents
        Do
            If TimerSpielerAnimation.Enabled = False Then Exit Do
            Sleep (MillisecSleep)
            DoEvents
        Loop
        blnMoveL�uft = False
        '------------------------------------------------------------
    ElseIf PlayerTurn = Links Then
        AusgespielteKarten(AKIndex).cType = GKType
        AusgespielteKarten(AKIndex).cValue = GKValue
        AusgespielteKarten(AKIndex).cGespieltVon = PlayerTurn
        StartpunktXLinks = (ScaleWidth / 2) - NormalWidth / 2 - 500                             'Gerbing 05.09.2017
        StartPunktYLinks = (ScaleHeight / 2) - NormalHeight / 2 - 100
        cdLinks(MyIndex).Visible = False                                                        'Gerbing 27.11.2017
        LinksGespielt.Width = NormalWidth
        LinksGespielt.Height = NormalHeight
        LinksGespielt.CardType = GKType
        LinksGespielt.CardValue = GKValue
        LinksGespielt.FaceMode = NachOben
        LinksGespielt.ZOrder                                                                    'Gerbing 27.09.2017
        blnMoveL�uft = True
        StartMillisek = timeGetTime
        
        Aufrufz�hlerLinks = 0
        TimerLinksAnimation.Enabled = True
        DoEvents
        Do
            If TimerLinksAnimation.Enabled = False Then Exit Do
            Sleep (MillisecSleep)
            DoEvents
        Loop
        blnMoveL�uft = False
        '------------------------------------------------------------
    ElseIf PlayerTurn = Mitte Then
        AusgespielteKarten(AKIndex).cType = GKType
        AusgespielteKarten(AKIndex).cValue = GKValue
        AusgespielteKarten(AKIndex).cGespieltVon = PlayerTurn
        StartpunktXMitte = (ScaleWidth / 2) - NormalWidth / 2                                   'Gerbing 05.09.2017
        StartPunktYMitte = cdMitte(MyIndex).Top
        cdMitte(MyIndex).Visible = False                                                        'Gerbing 27.11.2017
        MitteGespielt.Width = NormalWidth
        MitteGespielt.Height = NormalHeight
        MitteGespielt.CardType = GKType
        MitteGespielt.CardValue = GKValue
        MitteGespielt.FaceMode = NachOben
        MitteGespielt.ZOrder                                                                    'Gerbing 27.09.2017
        blnMoveL�uft = True
        
        Aufrufz�hlerMitte = 0
        TimerMitteAnimation.Enabled = True
        DoEvents
        Do
            If TimerMitteAnimation.Enabled = False Then Exit Do
            Sleep (MillisecSleep)
            DoEvents
        Loop
        blnMoveL�uft = False
        '------------------------------------------------------------
    ElseIf PlayerTurn = Rechts Then
        AusgespielteKarten(AKIndex).cType = GKType
        AusgespielteKarten(AKIndex).cValue = GKValue
        AusgespielteKarten(AKIndex).cGespieltVon = PlayerTurn
        StartpunktXRechts = (ScaleWidth / 2) - NormalWidth / 2 + 500                            'Gerbing 05.09.2017
        StartPunktYRechts = (ScaleHeight / 2) - NormalHeight / 2 - 100
        cdRechts(MyIndex).Visible = False                                                        'Gerbing 27.11.2017
        RechtsGespielt.Width = NormalWidth
        RechtsGespielt.Height = NormalHeight
        RechtsGespielt.CardType = GKType
        RechtsGespielt.CardValue = GKValue
        RechtsGespielt.FaceMode = NachOben
        RechtsGespielt.ZOrder                                                                   'Gerbing 27.09.2017
        blnMoveL�uft = True
        Aufrufz�hlerRechts = 0
        TimerRechtsAnimation.Enabled = True
        DoEvents
        Do
            If TimerRechtsAnimation.Enabled = False Then Exit Do
            Sleep (MillisecSleep)
            DoEvents
        Loop
        blnMoveL�uft = False
        '------------------------------------------------------------
    End If
End Sub

Private Function IsPlayable(Index As Integer) As Boolean
    Dim I As Integer
    Dim AI As Integer                                               'AusspielerIndex
    Dim blnFarbeVorhanden As Boolean
    
    'Kontrollieren ob der Spieler diese Karte spielen darf, abh�ngig vom gerade gew�hlte Spiel
    'bei Rot d�rfen die ersten zwei Karten keine roten Karten sein und es d�rfen keine roten Karten abgeworfen werden
    'ob der Spieler die geforderte Farbe hat oder ob er f�lschlicherweise eine Karte abwerfen will
    
    If PlayerTurn <> Spieler Then                               'Gerbing 08.09.2017
        Exit Function
    End If
    If Ausspieler <> Spieler Then
        'ein anderer hat ausgespielt
        'bei jeder Spielart muss zugegeben werden
        'frage ob der Spieler dieselbe Farbe hat wie der Ausspieler, aber eine andere Farbe spielen will
        If Ausspieler = Links Then AI = AKIndex - 2
        If Ausspieler = Mitte Then AI = AKIndex - 1
        If Ausspieler = Rechts Then AI = AKIndex
        
        For I = 1 To UBound(SpielerCards)
            If SpielerCards(I).cType = AusgespielteKarten(AI).cType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If cdSpieler(Index).CardType <> AusgespielteKarten(AI).cType And blnFarbeVorhanden = True Then
        'If cdSpieler(index).CardType <> AusgespielteKarten(AI).cType Then
            gstrMsg = "Du musst Farbe zugeben"
            frmMsgBox.Show 1
            IsPlayable = False
            'blnEingesammelt = True
            Exit Function
        End If
        If Spielart = Rot Then
            'es ist Rot
            If AKIndex < 8 Then                             'Die ersten zwei Stiche
                If cdSpieler(Index).CardType = 2 Then                  '2=Rot
                    gstrMsg = "Falsch gespielt bei Rot"
                    frmMsgBox.Show 1
                    IsPlayable = False
                    'blnEingesammelt = True
                    Exit Function
                End If
            End If
        End If
    Else
        'Der Ausspieler ist Spieler
        If Spielart = Rot Then
            'es ist Rot
            If AKIndex < 8 Then                             'Die ersten zwei Stiche
                If cdSpieler(Index).CardType = 2 Then                  '2=Rot
                    gstrMsg = "Falsch gespielt bei Rot"
                    frmMsgBox.Show 1
                    IsPlayable = False
                    'blnEingesammelt = True
                    Exit Function
                End If
            End If
        End If
    End If
    IsPlayable = True
End Function

Sub AlleKartenEntladen()
Dim I As Integer
    On Error Resume Next
    For I = cdLinks.Count To 1 Step -1
        Unload cdLinks.Item(I)
    Next I
    For I = cdMitte.Count To 1 Step -1
        Unload cdMitte.Item(I)
    Next I
    For I = cdRechts.Count To 1 Step -1
        Unload cdRechts.Item(I)
    Next I
    For I = cdSpieler.Count To 1 Step -1
        Unload cdSpieler.Item(I)
    Next I
    '-----------------------------------------
    For I = cdRoserlSchellA.Count To 1 Step -1
        Unload cdRoserlSchellA.Item(I)
    Next I
    For I = cdRoserlSchellB.Count To 1 Step -1
        Unload cdRoserlSchellB.Item(I)
    Next I
    For I = cdRoserlRotA.Count To 1 Step -1
        Unload cdRoserlRotA.Item(I)
    Next I
    For I = cdRoserlRotB.Count To 1 Step -1
        Unload cdRoserlRotB.Item(I)
    Next I
    For I = cdRoserlGr�nA.Count To 1 Step -1
        Unload cdRoserlGr�nA.Item(I)
    Next I
    For I = cdRoserlGr�nB.Count To 1 Step -1
        Unload cdRoserlGr�nB.Item(I)
    Next I
    For I = cdRoserlEichelA.Count To 1 Step -1
        Unload cdRoserlEichelA.Item(I)
    Next I
    For I = cdRoserlEichelB.Count To 1 Step -1
        Unload cdRoserlEichelB.Item(I)
    Next I
    '-----------------------------------------
    For I = cdRoserlSchellAss.Count To 1 Step -1
        Unload cdRoserlSchellAss.Item(I)
    Next I
    For I = cdRoserlRotAss.Count To 1 Step -1
        Unload cdRoserlRotAss.Item(I)
    Next I
    For I = cdRoserlGr�nAss.Count To 1 Step -1
        Unload cdRoserlGr�nAss.Item(I)
    Next I
    For I = cdRoserlEichelAss.Count To 1 Step -1
        Unload cdRoserlEichelAss.Item(I)
    Next I
    cdRoserlSchellA(0).Visible = False
    cdRoserlSchellB(0).Visible = False

    cdRoserlRotA(0).Visible = False
    cdRoserlRotB(0).Visible = False

    cdRoserlGr�nA(0).Visible = False
    cdRoserlGr�nB(0).Visible = False

    cdRoserlEichelA(0).Visible = False
    cdRoserlEichelB(0).Visible = False

    cdRoserlSchellAss(0).Visible = False
    cdRoserlRotAss(0).Visible = False
    cdRoserlGr�nAss(0).Visible = False
    cdRoserlEichelAss(0).Visible = False
    
    LinksGespielt.Visible = False                                                       'Gerbing 05.09.2017
    MitteGespielt.Visible = False
    RechtsGespielt.Visible = False
    SpielerGespielt.Visible = False
End Sub

Public Sub Kartengeben()
    Dim I As Integer
    
    On Error Resume Next
'    NewSuit = 0
    If blnSpielWiederholen = False Then
        If gblnNichtMischen = True Then
            Call NichtMischen
        Else                                                'Gerbing 10.12.2017
            Call Mischen                                    'Mischen
        End If
    End If
    Call AlleKartenEntladen
    ReDim LinksCards(1 To 8)
    ReDim MitteCards(1 To 8)
    ReDim RechtsCards(1 To 8)
    ReDim SpielerCards(1 To 8)
    
    For I = 0 To 7                                      'Gerbing 30.06.2017 Prototypfehler
'        LinksCards(i + 1) = CardDeck(i + 1)
'        MitteCards(i + 1) = CardDeck(i + 8)
'        RechtsCards(i + 1) = CardDeck(i + 16)
'        SpielerCards(i + 1) = CardDeck(i + 24)
        
        LinksCards(I + 1) = CardDeck(I + 1)
        MitteCards(I + 1) = CardDeck(I + 9)
        RechtsCards(I + 1) = CardDeck(I + 17)
        SpielerCards(I + 1) = CardDeck(I + 25)          'die 32.Karte hatte gar kein Spieler die wurde gesucht bei Lorum und nicht gefunden

        Load cdLinks(I)
        Load cdMitte(I)
        Load cdRechts(I)
        Load cdSpieler(I)
    Next I

    Call LinksKartenNeu
    Call MitteKartenNeu
    Call RechtsKartenNeu
    Call SpielerKartenNeu                               'Gerbing 20.09.2017 Spielerkarten zuletzt auflegen
End Sub

Sub MitteKartenNeu()                                    'Gerbing 28.07.2020
    Dim I As Integer
    Dim p As Integer
    Dim a As Integer
    Dim n As Integer
    Dim m As Integer
    Dim H�chsteFarbe As Integer
    Dim H�chsterWert As Integer

'    If cdMitte.Count = 8 Then
'        MitteStapelLeft = (ScaleWidth / 2) - (cdMitte(0).Width / 2 + ((cdMitte.Count - 1) * CARD_SPACE / 6) + CARD_SPACE / 6)   'Gerbing 26.11.2017
'    Else
'        MitteStapelLeft = (ScaleWidth / 2) - (cdMitte(0).Width / 2 + ((cdMitte.Count - 1) * CARD_SPACE / 6))                    'Gerbing 26.11.2017
'    End If
    'MitteStapelLeft = (ScaleWidth / 2) - ((cdSpieler.Count / 2) * CARD_SPACE) + 5                  'Fehler
    MitteStapelLeft = (ScaleWidth / 2) - ((cdMitte.Count / 2) * CARD_SPACE) + 5                     'Gerbing 22.10.2020
    MitteStapelTop = 50
    For I = 0 To cdMitte.Count
        On Error Resume Next
        cdMitte(I).Visible = False
        On Error GoTo 0
    Next I
    
        'Rangliste ermitteln
        ReDim Rangliste(UBound(MitteCards))
        On Error GoTo 0
        a = 1
        For I = 1 To UBound(MitteCards)
            H�chsteFarbe = 0
            H�chsterWert = 0
            For n = 1 To UBound(MitteCards)
                If MitteCards(n).cGespieltVon <> 9 Then
                    If Rangliste(n).Vergeben = False Then
                        If MitteCards(n).cType > H�chsteFarbe Then
                            H�chsteFarbe = MitteCards(n).cType
                            H�chsterWert = MitteCards(n).cValue
                        Else
                            If MitteCards(n).cType = H�chsteFarbe Then
                                If MitteCards(n).cValue >= H�chsterWert Then
                                    H�chsterWert = MitteCards(n).cValue
                                End If
                            End If
                        End If
                    End If
                End If
            Next n
            'Die Mittecards(m), die H�chsteFarbe und H�chsterWert enth�lt bekommt Rangliste(m).RangPosition = a und Rangliste(m).Vergeben = True
            For m = 1 To UBound(MitteCards)
                If MitteCards(m).cGespieltVon <> 9 Then
                    If Rangliste(m).Vergeben = False Then
                        If MitteCards(m).cType = H�chsteFarbe And MitteCards(m).cValue = H�chsterWert Then
                            Rangliste(m).RangPosition = a
                            Rangliste(m).Vergeben = True
                        End If
                    End If
                End If
            Next m
            a = a + 1
        Next I
        
    n = 0
    p = 0
    'Mittekarten entsprechend ihrer Rangliste auflegen
    For I = 1 To UBound(MitteCards)
        If MitteCards(I).cGespieltVon <> 9 Then
            On Error Resume Next
            cdMitte(n).FaceMode = NachOben
            cdMitte(n).Left = MitteStapelLeft + (CARD_SPACE * (Rangliste(I).RangPosition - 1))
            cdMitte(n).Top = MitteStapelTop
            cdMitte(n).Width = NormalWidth
            cdMitte(n).Height = NormalHeight
            If blnMaloiGetippt = True Then
                cdMitte(n).FaceMode = NachOben
            Else
                cdMitte(n).FaceMode = R�ckseite
            End If
            cdMitte(n).Visible = True
            'cdMitte(n).ZOrder                                                             'Gerbing 20.09.2017
            cdMitte(n).SetCardValue MitteCards(I).cType, MitteCards(I).cValue
            On Error GoTo 0
        Else
            p = p + 1
        End If
        n = n + 1
    Next I
End Sub

Sub SpielerKartenNeu()
    Dim I As Integer
    Dim p As Integer
    Dim a As Integer
    Dim n As Integer
    Dim m As Integer
    Dim H�chsteFarbe As Integer
    Dim H�chsterWert As Integer
    
    'SpielerStapelLeft = (ScaleWidth / 2) - (cdSpieler.Count / 2) * NormalWidth
    SpielerStapelLeft = (ScaleWidth / 2) - ((cdSpieler.Count / 2) * CARD_SPACE) + 5                 'Gerbing 26.11.2017
    SpielerStapelTop = ScaleHeight - 50 - NormalHeight
        'Rangliste ermitteln
        ReDim Rangliste(UBound(SpielerCards))
        On Error GoTo 0
        a = 1
        For I = 1 To UBound(SpielerCards)
            H�chsteFarbe = 0
            H�chsterWert = 0
            For n = 1 To UBound(SpielerCards)
                If SpielerCards(n).cGespieltVon <> 9 Then
                    If Rangliste(n).Vergeben = False Then
                        If SpielerCards(n).cType > H�chsteFarbe Then
                            H�chsteFarbe = SpielerCards(n).cType
                            H�chsterWert = SpielerCards(n).cValue
                        Else
                            If SpielerCards(n).cType = H�chsteFarbe Then
                                If SpielerCards(n).cValue >= H�chsterWert Then
                                    H�chsterWert = SpielerCards(n).cValue
                                End If
                            End If
                        End If
                    End If
                End If
            Next n
            'Die Spielercards(m), die H�chsteFarbe und H�chsterWert enth�lt bekommt Rangliste(m).RangPosition = a und Rangliste(m).Vergeben = True
            For m = 1 To UBound(SpielerCards)
                If SpielerCards(m).cGespieltVon <> 9 Then
                    If Rangliste(m).Vergeben = False Then
                        If SpielerCards(m).cType = H�chsteFarbe And SpielerCards(m).cValue = H�chsterWert Then
                            Rangliste(m).RangPosition = a
                            Rangliste(m).Vergeben = True
                        End If
                    End If
                End If
            Next m
            a = a + 1
        Next I
    
    For I = 0 To cdSpieler.Count
        On Error Resume Next
        cdSpieler(I).Visible = False
        cdSpieler(I).FaceMode = R�ckseite
        On Error GoTo 0
    Next I
    n = 0
    p = 0
    'Spielerkarten entsprechend ihrer Rangliste auflegen
    For I = 1 To UBound(SpielerCards)
        If SpielerCards(I).cGespieltVon <> 9 Then
            On Error Resume Next
            cdSpieler(n).FaceMode = NachOben
            cdSpieler(n).Left = SpielerStapelLeft + (CARD_SPACE * (Rangliste(I).RangPosition - 1))
            cdSpieler(n).Top = SpielerStapelTop
            cdSpieler(n).Width = NormalWidth
            cdSpieler(n).Height = NormalHeight
            cdSpieler(n).Visible = True
            cdSpieler(n).ZOrder                                                             'Gerbing 20.09.2017
            cdSpieler(n).SetCardValue SpielerCards(I).cType, SpielerCards(I).cValue
            On Error GoTo 0
        Else
            p = p + 1
        End If
        n = n + 1
    Next I
End Sub

Sub LinksKartenNeu()
    Dim I As Integer
    Dim p As Integer
    Dim a As Integer
    Dim n As Integer
    Dim m As Integer
    Dim H�chsteFarbe As Integer
    Dim H�chsterWert As Integer
    Dim BesterRang As Integer                                                           'Der beste Rang ist der niedrigste Wert

    LinksStapelLeft = 50
    LinksStapelTop = (ScaleHeight / 2) - NormalHeight                                   'Gerbing 20.09.2017
        'Rangliste ermitteln
        ReDim Rangliste(UBound(LinksCards))
        On Error GoTo 0
        a = 1
        For I = 1 To UBound(LinksCards)
            H�chsteFarbe = 0
            H�chsterWert = 0
            For n = 1 To UBound(LinksCards)
                If LinksCards(n).cGespieltVon <> 9 Then
                    If Rangliste(n).Vergeben = False Then
                        If LinksCards(n).cType > H�chsteFarbe Then
                            H�chsteFarbe = LinksCards(n).cType
                            H�chsterWert = LinksCards(n).cValue
                        Else
                            If LinksCards(n).cType = H�chsteFarbe Then
                                If LinksCards(n).cValue >= H�chsterWert Then
                                    H�chsterWert = LinksCards(n).cValue
                                End If
                            End If
                        End If
                    End If
                End If
            Next n
            'Die Linkscards(m), die H�chsteFarbe und H�chsterWert enth�lt bekommt Rangliste(m).RangPosition = a und Rangliste(m).Vergeben = True
            For m = 1 To UBound(LinksCards)
                If LinksCards(m).cGespieltVon <> 9 Then
                    If Rangliste(m).Vergeben = False Then
                        If LinksCards(m).cType = H�chsteFarbe And LinksCards(m).cValue = H�chsterWert Then
                            Rangliste(m).RangPosition = a
                            Rangliste(m).Vergeben = True
                        End If
                    End If
                End If
            Next m
            a = a + 1
        Next I
    For I = 0 To cdLinks.Count
        On Error Resume Next
        cdLinks(I).Visible = False
        On Error GoTo 0
    Next I
    'Der Rangliste entsprechend nach unten auflegen Gerbing 28.07.2020
    'Der beste Rang ist der niedrigste Wert
    For p = 1 To UBound(LinksCards)
        BesterRang = 9
        For I = 1 To UBound(LinksCards)
            If Rangliste(I).RangPosition <> 9 Then
                If Rangliste(I).RangPosition < BesterRang Or BesterRang = 8 Then
                    BesterRang = Rangliste(I).RangPosition
                    n = I
                End If
            End If
        Next I
        Rangliste(n).RangPosition = 9                                           '9 ist das Kennzeichen f�r schon gefunden
        On Error Resume Next
        If blnMaloiGetippt = True Then
            cdLinks(p - 1).FaceMode = NachOben
        Else
            cdLinks(p - 1).FaceMode = R�ckseite
        End If
        '-----------------------------------
        cdLinks(p - 1).Left = LinksStapelLeft                                     'Gerbing 20.09.2017
        cdLinks(p - 1).Top = LinksStapelTop + (CARD_SPACE / 3 * (p - 1))
        cdLinks(p - 1).Width = NormalWidth
        cdLinks(p - 1).Height = NormalHeight
        cdLinks(p - 1).Visible = True
        cdLinks(p - 1).SetCardValue LinksCards(n).cType, LinksCards(n).cValue
        On Error GoTo 0
    Next p
End Sub

Sub RechtsKartenNeu()
    Dim I As Integer
    Dim p As Integer
    Dim a As Integer
    Dim n As Integer
    Dim m As Integer
    Dim H�chsteFarbe As Integer
    Dim H�chsterWert As Integer
    Dim BesterRang As Integer                                                           'Der beste Rang ist der niedrigste Wert

    RechtsStapelLeft = ScaleWidth - 50 - CARD_SPACE
    RechtsStapelTop = (ScaleHeight / 2) - NormalHeight                                   'Gerbing 20.09.2017
        'Rangliste ermitteln
        ReDim Rangliste(UBound(RechtsCards))
        On Error GoTo 0
        a = 1
        For I = 1 To UBound(RechtsCards)
            H�chsteFarbe = 0
            H�chsterWert = 0
            For n = 1 To UBound(RechtsCards)
                If RechtsCards(n).cGespieltVon <> 9 Then
                    If Rangliste(n).Vergeben = False Then
                        If RechtsCards(n).cType > H�chsteFarbe Then
                            H�chsteFarbe = RechtsCards(n).cType
                            H�chsterWert = RechtsCards(n).cValue
                        Else
                            If RechtsCards(n).cType = H�chsteFarbe Then
                                If RechtsCards(n).cValue >= H�chsterWert Then
                                    H�chsterWert = RechtsCards(n).cValue
                                End If
                            End If
                        End If
                    End If
                End If
            Next n
            'Die rechtscards(m), die H�chsteFarbe und H�chsterWert enth�lt bekommt Rangliste(m).RangPosition = a und Rangliste(m).Vergeben = True
            For m = 1 To UBound(RechtsCards)
                If RechtsCards(m).cGespieltVon <> 9 Then
                    If Rangliste(m).Vergeben = False Then
                        If RechtsCards(m).cType = H�chsteFarbe And RechtsCards(m).cValue = H�chsterWert Then
                            Rangliste(m).RangPosition = a
                            Rangliste(m).Vergeben = True
                        End If
                    End If
                End If
            Next m
            a = a + 1
        Next I
    For I = 0 To cdRechts.Count
        On Error Resume Next
        cdRechts(I).Visible = False
        On Error GoTo 0
    Next I
    'Der Rangliste entsprechend nach unten auflegen Gerbing 28.07.2020
    'Der beste Rang ist der niedrigste Wert
    For p = 1 To UBound(RechtsCards)
        BesterRang = 9
        For I = 1 To UBound(RechtsCards)
            If Rangliste(I).RangPosition <> 9 Then
                If Rangliste(I).RangPosition < BesterRang Or BesterRang = 8 Then
                    BesterRang = Rangliste(I).RangPosition
                    n = I
                End If
            End If
        Next I
        Rangliste(n).RangPosition = 9                                           '9 ist das Kennzeichen f�r schon gefunden
        On Error Resume Next
        If blnMaloiGetippt = True Then
            cdRechts(p - 1).FaceMode = NachOben
        Else
            cdRechts(p - 1).FaceMode = R�ckseite
        End If
        '-----------------------------------
        cdRechts(p - 1).Left = RechtsStapelLeft                                     'Gerbing 20.09.2017
        cdRechts(p - 1).Top = RechtsStapelTop + (CARD_SPACE / 3 * (p - 1))
        cdRechts(p - 1).Width = NormalWidth
        cdRechts(p - 1).Height = NormalHeight
        cdRechts(p - 1).Visible = True
        cdRechts(p - 1).SetCardValue RechtsCards(n).cType, RechtsCards(n).cValue
        On Error GoTo 0
    Next p
End Sub

Private Sub Form_Initialize()
    Set IniFso = New FileSystemObject
    Set LogFso = New FileSystemObject
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Eintippen maloi zeigt die Karten des Computers
    Static K1 As String, K2 As String, K3 As String, K4 As String, K5 As String
    
    Dim ShiftDown As Boolean
    Dim AltDown As Boolean
    Dim CtrlDown As Boolean
    
    K5 = K4
    K4 = K3
    K3 = K2
    K2 = K1
    K1 = UCase$(Chr$(KeyCode))
    ' Cheat code.
    If K5 = "M" And K4 = "A" And K3 = "L" And K2 = "O" And K1 = "I" Then
        blnMaloiGetippt = True
        Call ComputerKartenAufdecken
    End If
    
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0

    Select Case KeyCode
        Case vbKeyReturn                                                        'Gerbing 09.12.2017
            Call Form_Click                                                     'Gerbing 09.12.2017
        Case vbKeyN
            If CtrlDown Then                               'Strg+N gleichzeitig
                'frmSpielWiederholung.Show 1
                gblnMsgBoxAntwort = True
                gstrMsg = "Du kannst dieses Spiel noch einmal mit den gleichen Karten spielen." & vbNewLine
                gstrMsg = gstrMsg & "Willst Du sofort das neue Spiel beginnen, w�hle Ja" & vbNewLine
                gstrMsg = gstrMsg & "Willst Du das aktuelle Spiel erst noch zu Ende spielen, w�hle Nein"
                frmMsgBox.Show 1
                If glngAntwort = vbYes Then
                    blnSpielWiederholen = True
                    Call Spielende
                Else
                    blnSpielWiederholen = True
                End If
            End If
        Case vbKeyS
            If CtrlDown Then
                frmSpielBogen.Show
            End If
        Case vbKeyE
            gblnMsgBoxAntwort = True
            gstrMsg = "Willst Du das Spiel abbrechen?"
            frmMsgBox.Show 1
            If glngAntwort = vbYes Then
                End
            End If
    End Select
End Sub

Sub ComputerKartenAufdecken()
    Call LinksKartenNeu
    Call MitteKartenNeu
    Call RechtsKartenNeu
End Sub

Private Sub Form_Load()
    Dim I As Integer
    
    #If Spielwiederholen = 0 Then                   '0=nein                                 'Gerbing 19.05.2021
        Me.Caption = "RoserlO"
    #Else
        Me.Caption = "RoserlM"
    #End If
    
    Call ReadRoserlIniFile
    Call MindestAufl�sung
    Call AnpassenNutzerWunsch(Me)
    For I = 0 To 5
        lblRoserl(I).FontSize = 20
    Next I
    blnSchellZEM = True
    blnRotZEM = True
    blnGr�nZEM = True
    blnEichelZEM = True
    MillisecSleep = 5                                                                       'Gerbing 09.11.2017
    MillisecSleepLorum = 20
    iMaxMove = 50
    iMulti = 2
    TimerSpielerAnimation.Interval = MillisecSleep                                          'Gerbing 09.11.2017
    TimerLinksAnimation.Interval = MillisecSleep                                            'Gerbing 09.11.2017
    TimerMitteAnimation.Interval = MillisecSleep                                            'Gerbing 09.11.2017
    TimerRechtsAnimation.Interval = MillisecSleep                                           'Gerbing 09.11.2017

    frmSpielBogen.lblDatum = Date
    Call AnpassenNutzerWunsch(frmSpielBogen)
End Sub

Private Sub Form_Resize()
    Dim n As Integer
    Dim I As Integer
    
    On Error Resume Next
    If Me.Width < MIN_DIMEN Then Width = MIN_DIMEN
    If Me.Height < MIN_DIMEN Then Height = MIN_DIMEN
    On Error GoTo 0
    lblMitte.Top = 25
    lblSpieler.Top = ScaleHeight - 25
    lblMitte.Left = ScaleWidth / 2 - 25
    lblSpieler.Left = ScaleWidth / 2 - 25
    lblRechts.Top = 25
    lblRechts.Left = ScaleWidth - 150                                                       'Gerbing 26.11.2017
    lblLinks.Top = 25
    lblLinks.Left = 100                                                                     'Gerbing 26.11.2017
    lblLinks.Visible = True
    lblMitte.Visible = True
    lblRechts.Visible = True
    lblSpieler.Visible = True
    n = 3
    For I = 0 To 5
        lblRoserl(I).Left = (ScaleWidth / 2) - (n * lblRoserl(I).Width)
        n = n - 1
    Next I
    lblRunde.Left = lblRoserl(0).Left
    lblRunde.Top = lblRoserl(0).Top + lblRoserl(0).Height
    lblRunde.Width = lblRoserl(0).Width * 6
    On Error Resume Next
    I = UBound(LinksCards)
    If Err.Number <> 0 Then GoTo ExitSub
    If Spielart <> Lorum Then                                                               'Gerbing 19.09.2017
        If UBound(LinksCards) <> 0 Then
            Call LinksKartenNeu
            Call MitteKartenNeu
            Call RechtsKartenNeu
            Call SpielerKartenNeu
            MitteGespielt.Left = (ScaleWidth / 2) - CARD_WIDTH / 2
            MitteGespielt.Width = NormalWidth
            MitteGespielt.Height = NormalHeight
            LinksGespielt.Left = 100
            LinksGespielt.Width = NormalWidth
            LinksGespielt.Height = NormalHeight
            RechtsGespielt.Left = ScaleWidth - CARD_WIDTH - 100
            RechtsGespielt.Width = NormalWidth
            RechtsGespielt.Height = NormalHeight
            SpielerGespielt.Left = (ScaleWidth / 2) - CARD_WIDTH / 2
            SpielerGespielt.Top = cdSpieler(0).Top - CARD_HEIGHT / 4
            SpielerGespielt.Width = NormalWidth
            SpielerGespielt.Height = NormalHeight
        End If
    End If
ExitSub:
End Sub

Private Sub Einsammeln()
    Dim Einsammler As Integer
    Dim strEinsammler As String
    Dim H�chsteKarte As Integer
    Dim I As Integer

    'Einsammeln kann nur ein Mitspieler, der eine h�here Karte der gleichen Farbe gespielt hat als der Ausspieler
    'Wenn kein Mitspieler die gleiche Farbe gespielt hat (also abgeworfen) dann ist der Ausspieler der Einsammler
    'Nach dem Einsammeln steht fest, wer als n�chster ausspielen muss
    'und ich Kann kennzeichnen wer die Karten eingesammelt hat

    
    If Ausspieler = Rechts Then
        Einsammler = Rechts                                     'Wenn alle abwerfen bleibt rechts der Einsammler
        H�chsteKarte = RechtsGespielt.CardValue
        If SpielerGespielt.CardType = RechtsGespielt.CardType And SpielerGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = SpielerGespielt.CardValue
            Einsammler = Spieler
        End If
        If LinksGespielt.CardType = RechtsGespielt.CardType And LinksGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = LinksGespielt.CardValue
            Einsammler = Links
        End If
        If MitteGespielt.CardType = RechtsGespielt.CardType And MitteGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = MitteGespielt.CardValue
            Einsammler = Mitte
        End If
    End If
    If Ausspieler = Spieler Then
        Einsammler = Spieler                                    'Wenn alle abwerfen bleibt spieler der Einsammler
        H�chsteKarte = SpielerGespielt.CardValue
        If LinksGespielt.CardType = SpielerGespielt.CardType And LinksGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = LinksGespielt.CardValue
            Einsammler = Links
        End If
        If MitteGespielt.CardType = SpielerGespielt.CardType And MitteGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = MitteGespielt.CardValue
            Einsammler = Mitte
        End If
        If RechtsGespielt.CardType = SpielerGespielt.CardType And RechtsGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = RechtsGespielt.CardValue
            Einsammler = Rechts
        End If
    End If
    If Ausspieler = Links Then
        Einsammler = Links                                    'Wenn alle abwerfen bleibt links der Einsammler
        H�chsteKarte = LinksGespielt.CardValue
        If MitteGespielt.CardType = LinksGespielt.CardType And MitteGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = MitteGespielt.CardValue
            Einsammler = Mitte
        End If
        If RechtsGespielt.CardType = LinksGespielt.CardType And RechtsGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = RechtsGespielt.CardValue
            Einsammler = Rechts
        End If
        If SpielerGespielt.CardType = LinksGespielt.CardType And SpielerGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = SpielerGespielt.CardValue
            Einsammler = Spieler
        End If
    End If
    If Ausspieler = Mitte Then
        Einsammler = Mitte                                    'Wenn alle abwerfen bleibt mitte der Einsammler
        H�chsteKarte = MitteGespielt.CardValue
        If RechtsGespielt.CardType = MitteGespielt.CardType And RechtsGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = RechtsGespielt.CardValue
            Einsammler = Rechts
        End If
        If SpielerGespielt.CardType = MitteGespielt.CardType And SpielerGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = SpielerGespielt.CardValue
            Einsammler = Spieler
        End If
        If LinksGespielt.CardType = MitteGespielt.CardType And LinksGespielt.CardValue > H�chsteKarte Then
            H�chsteKarte = LinksGespielt.CardValue
            Einsammler = Links
        End If
    End If
    '--------------------------------------------------------------------------------------------------------
    'kennzeichnen wer die Karten eingesammelt hat und Kontrolle ob alle Ober, alle Roten, RoterK�nig ausgespielt sind
    blnRKGespielt = False
    For I = AKIndex - 3 To AKIndex
        AusgespielteKarten(I).cEingesammeltVon = Einsammler     'Kennzeichnen wer die Karten eingesammelt hat
        If AusgespielteKarten(I).cType = 2 And AusgespielteKarten(I).cValue = 7 And Spielart = RoterK�nig Then
            blnRKGespielt = True                                'Roter K�nig gespielt
        End If
        If AusgespielteKarten(I).cValue = 6 And Spielart = Ober Then    'Gerbing 12.09.2017
            Z�hlerOber = Z�hlerOber + 1
            If Z�hlerOber = 4 Then
                blnAlleOberGespielt = True
            End If
        End If
        If AusgespielteKarten(I).cType = 2 And Spielart = Rot Then      'Gerbing 12.09.2017
            Z�hlerRote = Z�hlerRote + 1
            If Z�hlerRote = 8 Then
                blnAlleRotenGespielt = True
            End If
        End If
    Next I
    '--------------------------------------------------------------------------------------------------------
    Select Case Einsammler
        Case 0
            strEinsammler = "Spieler"
            lblSpieler.Font.Bold = True                         'Gerbing 15.08.2017
            lblSpieler.ForeColor = FarbeAusspieler                         'schwarz
            lblLinks.ForeColor = &HFFFFFF                       'wei�
            lblMitte.ForeColor = &HFFFFFF
            lblRechts.ForeColor = &HFFFFFF
            lblLinks.Font.Bold = False
            lblMitte.Font.Bold = False
            lblRechts.Font.Bold = False
        Case 1
            strEinsammler = "Links"
            lblLinks.Font.Bold = True                           'Gerbing 15.08.2017
            lblLinks.ForeColor = FarbeAusspieler                           'schwarz
            lblSpieler.ForeColor = &HFFFFFF                     'wei�
            lblMitte.ForeColor = &HFFFFFF
            lblRechts.ForeColor = &HFFFFFF
            lblMitte.Font.Bold = False
            lblRechts.Font.Bold = False
            lblSpieler.Font.Bold = False
        Case 2
            strEinsammler = "Mitte"
            lblMitte.Font.Bold = True                           'Gerbing 15.08.2017
            lblMitte.ForeColor = FarbeAusspieler                           'schwarz
            lblLinks.ForeColor = &HFFFFFF                       'wei�
            lblSpieler.ForeColor = &HFFFFFF
            lblRechts.ForeColor = &HFFFFFF
            lblLinks.Font.Bold = False
            lblRechts.Font.Bold = False
            lblSpieler.Font.Bold = False
        Case 3
            strEinsammler = "Rechts"
            lblRechts.Font.Bold = True                          'Gerbing 15.08.2017
            lblRechts.ForeColor = FarbeAusspieler                          'schwarz
            lblLinks.ForeColor = &HFFFFFF                       'wei�
            lblMitte.ForeColor = &HFFFFFF
            lblSpieler.ForeColor = &HFFFFFF
            lblLinks.Font.Bold = False
            lblMitte.Font.Bold = False
            lblSpieler.Font.Bold = False
    End Select

'    gstrMsg = "Der Stich geht an " & strEinsammler
'    frmMsgBox.Show 1

    'Anstelle von "Der Stich geht an... " will ich gleichzeitige Bewegungsanimation in Richtung Einsammler machen
    If Einsammler = Links Then
        blnMoveL�uft = True
        Aufrufz�hlerEinsammler = 0
        TimerEinsammlerLinks.Enabled = True
        DoEvents
        Do
            If TimerEinsammlerLinks.Enabled = False Then Exit Do
            Sleep (MillisecSleep)
            DoEvents
        Loop
        blnMoveL�uft = False
    ElseIf Einsammler = Mitte Then
        blnMoveL�uft = True
        Aufrufz�hlerEinsammler = 0
        TimerEinsammlerMitte.Enabled = True
        DoEvents
        Do
            If TimerEinsammlerMitte.Enabled = False Then Exit Do
            Sleep (MillisecSleep)
            DoEvents
        Loop
        blnMoveL�uft = False
    ElseIf Einsammler = Rechts Then
        blnMoveL�uft = True
        Aufrufz�hlerEinsammler = 0
        TimerEinsammlerRechts.Enabled = True
        DoEvents
        Do
            If TimerEinsammlerRechts.Enabled = False Then Exit Do
            Sleep (MillisecSleep)
            DoEvents
        Loop
        blnMoveL�uft = False
    ElseIf Einsammler = Spieler Then
        blnMoveL�uft = True
        Aufrufz�hlerEinsammler = 0
        TimerEinsammlerSpieler.Enabled = True
        DoEvents
        Do
            If TimerEinsammlerSpieler.Enabled = False Then Exit Do
            Sleep (MillisecSleep)
            DoEvents
        Loop
        blnMoveL�uft = False
        'SchlafTimer.Enabled = True                                                      'Gerbing 20.09.2017 25.09.2017
    End If
    '-----------------------------------------------------------------
    Ausspieler = Einsammler
    'Nach dem Einsammeln steht fest, wer als n�chster ausspielen muss
    blnEingesammelt = True                                      'Gerbing 13.08.2017
    blnSpielerWarDran = False
    blnLinksWarDran = False
    blnMitteWarDran = False
    blnRechtsWarDran = False
    
    SpielerGespielt.Visible = False
    LinksGespielt.Visible = False
    MitteGespielt.Visible = False
    RechtsGespielt.Visible = False
End Sub

Private Sub Spielende()
    Dim strAusspieler As String
    
    'Soll das Spiel wiederholt werden weil Strg+N gedr�ckt wurde?
    If blnSpielWiederholen = True Then
        gstrMsg = "Das Spiel wird wiederholt"
        frmMsgBox.Show 1
        Call Kartengeben
        blnAlleOberGespielt = False                                         'Gerbing 29.11.2017
        blnRKGespielt = False                                               'Gerbing 29.11.2017
        blnAlleRotenGespielt = False                                        'Gerbing 29.11.2017
        blnLorumSpielende = False                                           'Gerbing 05.09.2017
        blnSpielWiederholen = False                                         'Gerbing 05.09.2017
        Ausspieler = AusspielerBeiSpielbeginn                               'Gerbing 05.09.2017
        AKIndex = 0                                                         'Gerbing 05.09.2017
        Call AusspielerFett                                                 'Gerbing 05.09.2017
        blnEingesammelt = True
        If Spielart = Lorum Then
            PlayerTurnRoserl = Ausspieler                                   'Gerbing 05.09.2017
        End If
        Exit Sub
    End If
    '-------------------------------------------------------
    'das Spiel ist zu Ende
    Call PunkteZ�hlen
    'Soll das Spiel wiederholt werden weil Button 'Spiel wiederholen in frmSpielEnde gedr�ckt wurde? Gerbing 26.07.2020
    If blnSpielWiederholen = True Then
        gstrMsg = "Das Spiel wird wiederholt"
        frmMsgBox.Show 1
        Call Kartengeben
        blnAlleOberGespielt = False
        blnRKGespielt = False
        blnAlleRotenGespielt = False
        blnLorumSpielende = False
        blnSpielWiederholen = False
        Ausspieler = AusspielerBeiSpielbeginn
        AKIndex = 0
        Call AusspielerFett
        blnEingesammelt = True
        If Spielart = Lorum Then
            PlayerTurnRoserl = Ausspieler
        End If
        Exit Sub
    End If
    '-------------------------------------------------------
    blnMaloiGetippt = False
    'Call Spielabrechnung
    AKIndex = 0
    AusspielerBeiSpielbeginn = AusspielerBeiSpielbeginn + 1
    If AusspielerBeiSpielbeginn = 4 Then
        AusspielerBeiSpielbeginn = AusspielerBeiSpielbeginn - 4
    End If
    Ausspieler = AusspielerBeiSpielbeginn
    If Spielart <> Lorum Then
        'Wenn nicht Lorum dann keine Spielabrechnung
        frmRoserl.lblRoserl(Spielart).BackColor = &H8000000F                'Farbe zur�cksetzen
        Spielart = Spielart + 1                                             'n�chste Spielart
        frmRoserl.lblRoserl(Spielart).BackColor = &HFF&                     'Buchstabe f�r das n�chste Spiel wird rot
        If Spielart = Lorum Then                                            'wird das n�chste Spiel=Lorum ?
            ' das n�chste Spiel wird Lorum
            'blnMaloiGetippt = True                                          'Gerbing 11.07.2017 Lorum kann mit aufgedeckten Karten gespielt werden
            If Ausspieler = Links Then
                strAusspieler = "Links"
                lblLinks.Font.Bold = True                           'Gerbing 15.08.2017
                lblLinks.ForeColor = FarbeAusspieler                           'schwarz
                lblSpieler.ForeColor = &HFFFFFF                     'wei�
                lblMitte.ForeColor = &HFFFFFF
                lblRechts.ForeColor = &HFFFFFF
                lblMitte.Font.Bold = False
                lblRechts.Font.Bold = False
                lblSpieler.Font.Bold = False
            End If
            If Ausspieler = Mitte Then
                strAusspieler = "Mitte"
                lblMitte.Font.Bold = True                           'Gerbing 15.08.2017
                lblMitte.ForeColor = FarbeAusspieler                           'schwarz
                lblLinks.ForeColor = &HFFFFFF                       'wei�
                lblSpieler.ForeColor = &HFFFFFF
                lblRechts.ForeColor = &HFFFFFF
                lblLinks.Font.Bold = False
                lblRechts.Font.Bold = False
                lblSpieler.Font.Bold = False
            End If
            If Ausspieler = Rechts Then
                strAusspieler = "Rechts"
                lblRechts.Font.Bold = True                          'Gerbing 15.08.2017
                lblRechts.ForeColor = FarbeAusspieler                          'schwarz
                lblLinks.ForeColor = &HFFFFFF                       'wei�
                lblMitte.ForeColor = &HFFFFFF
                lblSpieler.ForeColor = &HFFFFFF
                lblLinks.Font.Bold = False
                lblMitte.Font.Bold = False
                lblSpieler.Font.Bold = False
            End If
            If Ausspieler = Spieler Then
                strAusspieler = "Spieler"
                lblSpieler.Font.Bold = True                         'Gerbing 15.08.2017
                lblSpieler.ForeColor = FarbeAusspieler                         'schwarz
                lblLinks.ForeColor = &HFFFFFF                       'wei�
                lblMitte.ForeColor = &HFFFFFF
                lblRechts.ForeColor = &HFFFFFF
                lblLinks.Font.Bold = False
                lblMitte.Font.Bold = False
                lblRechts.Font.Bold = False
            End If
            Call StandardWerteLorum
            PlayerTurnRoserl = Ausspieler
'            gstrMsg = "Der Lorum-Ausspieler ist " & strAusspieler  'Gerbing 14.10.2017
'            frmMsgBox.Show 1
        End If
    Else
        'nach Lorum kommt die Spielabrechnung
        Call Spielabrechnung
        If Rundez�hler <> 4 Then
'            gblnMsgBoxAntwort = True
'            gstrMsg = "Es wird viermal gespielt, willst Du weiter spielen?"
'            gstrMsg = "Die n�chste Spielrunde beginnt"
'            frmMsgBox.Show 1
'            If glngAntwort = vbNo Then
'                AlleKartenEntladen
'                Unload Me
'                End
'            End If
            'Rundez�hler erh�hen                                        'Gerbing 18.09.2017
            If Rundez�hler = 1 Then
                lblRunde.Caption = "2. Runde"
            ElseIf Rundez�hler = 2 Then
                lblRunde.Caption = "3. Runde"
            ElseIf Rundez�hler = 3 Then
                lblRunde.Caption = "4. Runde"
            End If
            Rundez�hler = Rundez�hler + 1
        Else
            'Rundez�hler = 4
            Call EintragenBestwerte                                         'Gerbing 31.07.2017
            gblnMsgBoxAntwort = True
            gstrMsg = "Es sind 4 Runden gespielt. Das Spiel ist zu Ende. Willst Du weiter spielen?"    'Gerbing 17.08.2017
            frmMsgBox.Show 1
            If glngAntwort = vbNo Then
                AlleKartenEntladen
                Unload Me
                End
            Else
                Rundez�hler = 1                                         'Gerbing 17.08.2017
                lblRunde.Caption = "1. Runde"                       'Gerbing 28.09.2017
                'Spielbogen muss auf Null zur�ckgesetzt werden              'Gerbing 03.09.2017
                Call SpielBogenNull
            End If
        End If
        'AusspielerBeiSpielbeginn um eins zur�cksetzen damit jeder einmal bei jedem Spiel der Ausspieler ist
        If AusspielerBeiSpielbeginn = 0 Then
            AusspielerBeiSpielbeginn = 3
        Else
            AusspielerBeiSpielbeginn = AusspielerBeiSpielbeginn - 1
        End If
        frmRoserl.lblRoserl(Spielart).BackColor = &H8000000F                'Farbe zur�cksetzen
        Spielart = Rot                                                      'erstes Spiel ist Spielart Rot
        frmRoserl.lblRoserl(Rot).BackColor = &HFF&                          'beim ersten Spiel Buchstabe R wird rot
        Call StandardWerteNormal
    End If
    '---------------------------------------------
    blnLorumSpielende = False
    Erase SpielerCards
    Erase LinksCards
    Erase MitteCards
    Erase RechtsCards
    
    Call Kartengeben
    Ausspieler = AusspielerBeiSpielbeginn
    PlayerTurn = Ausspieler
    blnEingesammelt = True                                                  'Gerbing 13.08.2017
    LinksGespielt.Visible = False
    MitteGespielt.Visible = False
    RechtsGespielt.Visible = False
    SpielerGespielt.Visible = False
    blnNeuesSpielBeginnt = True
    blnSchellZEM = True
    blnRotZEM = True
    blnGr�nZEM = True
    blnEichelZEM = True
    If Spielart = Rot Then                                                  'Gerbing 14.10.2017
        gstrMsg = "Die n�chste Spielrunde beginnt." & vbNewLine             'Gerbing 19.12.2017
        gstrMsg = gstrMsg & "Du musst zum Weiterspielen auf den roten Buchstabe klicken" & vbNewLine
        frmMsgBox.Show 1
    End If
End Sub

Private Sub Spielabrechnung()
    On Error Resume Next
    frmSpielBogen.Show 1                                            'Gerbing 18.11.2017
    frmSpielBogen.ZOrder                                            'Gerbing 19.12.2017
    On Error GoTo 0
End Sub

Private Function SucheBesteKarte() As Integer
    Dim I As Integer
    Dim NiedrigstePosition As Integer
    
    'Hier wird die beste Karte gesucht, wenn der Computer eine Karte spielt
    'Es ist zu unterscheiden, ob der Computer der Ausspieler ist, oder ob er eine Karte zugeben muss
    
    If Ausspieler = PlayerTurn Then
        'Der Computer ist der Ausspieler
        'If Spielart = Rot And AKIndex < 9 Then                     'Gerbing 14.10.2017
        If Spielart = Rot And AKIndex < 8 Then                      'Gerbing 14.10.2017
            'Spielart = Rot erster oder zweiter Stich
            'Durchsuche alle Karten au�er rote bei den ersten zwei Stichen und finde die niedrigste Zum Ausspielen
            SucheBesteKarte = SucheBesteKarteRotOderSticheAusspieler
            'Gerbing 15.10.2017---------------------------------------Start
            If Ausspieler = Links Then
                If LinksCards(SucheBesteKarte).cType = 1 Then         '1=Schell
                    blnSchellZEM = False
                ElseIf LinksCards(SucheBesteKarte).cType = 2 Then     '2=Rot
                    blnRotZEM = False
                ElseIf LinksCards(SucheBesteKarte).cType = 3 Then     '3=Gr�n
                    blnGr�nZEM = False
                ElseIf LinksCards(SucheBesteKarte).cType = 4 Then     '4=Eichel
                    blnEichelZEM = False
                End If
            ElseIf Ausspieler = Mitte Then
                If MitteCards(SucheBesteKarte).cType = 1 Then         '1=Schell
                    blnSchellZEM = False
                ElseIf MitteCards(SucheBesteKarte).cType = 2 Then     '2=Rot
                    blnRotZEM = False
                ElseIf MitteCards(SucheBesteKarte).cType = 3 Then     '3=Gr�n
                    blnGr�nZEM = False
                ElseIf MitteCards(SucheBesteKarte).cType = 4 Then     '4=Eichel
                    blnEichelZEM = False
                End If
            ElseIf Ausspieler = Rechts Then
                If RechtsCards(SucheBesteKarte).cType = 1 Then         '1=Schell
                    blnSchellZEM = False
                ElseIf RechtsCards(SucheBesteKarte).cType = 2 Then     '2=Rot
                    blnRotZEM = False
                ElseIf RechtsCards(SucheBesteKarte).cType = 3 Then     '3=Gr�n
                    blnGr�nZEM = False
                ElseIf RechtsCards(SucheBesteKarte).cType = 4 Then     '4=Eichel
                    blnEichelZEM = False
                End If
            End If
            'Gerbing 15.10.2017---------------------------------------End
            Exit Function
        End If
        '---------------------------------------------------------------------------------------------------
        If Spielart = ErsterUndLetzterstich Then
            'Hier darf der Computer auch hohe Karten ausspielen, wenn es nicht der ErsterUndLetzterStich ist
            SucheBesteKarte = SucheBesteKarteErsterUndLetzter
            Exit Function
        End If
        '---------------------------------------------------------------------------------------------------
        If Spielart = Ober Then
            'Wenn abgeworfen werden kann, zuerst einen Ober abwerfen,
            'wenn zugegeben werden muss, den Ober nur wenn er der letzte dieser Farbe ist
            'wenn ausgespielt werden muss, den Ober nur wenn er der letzte dieser Farbe ist
            SucheBesteKarte = SucheBesteKarteOberAusspieler
            Exit Function
        End If
        '---------------------------------------------------------------------------------------------------
        If Spielart = RoterK�nig Then
            'nach dem Ausspielen von RoterK�nig kann das Spiel beendet werden
            'auf keinen Fall den roten K�nig ausspielen, wenn es noch andere rote Karten gibt
            SucheBesteKarte = SucheBesteKarteRoterK�nigAusspieler
            Exit Function
        End If
        '---------------------------------------------------------------------------------------------------
        If Spielart = Lorum Then
            'Lorum hat eigene Regeln
            SucheBesteKarte = SucheBesteKarteLorumAusspieler
            Exit Function
        End If
        '---------------------------------------------------------------------------------------------------
        'Hier ist es Spielart (Rot nach dem zweiten Stich) oder Stiche
        'Durchsuche alle Karten und finde die niedrigste Zum Ausspielen
        'Es ist sinnlos eine Farbe auszuspielen, die nur nur noch der Computer (Links Mitte Rechts) selbst hat
        SucheBesteKarte = SucheBesteKarteRotOderSticheAusspieler_N2
    Else
        'Der computer spielt nicht aus sondern muss zugeben oder abwerfen
        If Spielart = Ober Then
            'Wenn abgeworfen werden kann, zuerst einen Ober abwerfen,
            'wenn zugegeben werden muss, den Ober nur wenn er der letzte dieser Farbe ist
            SucheBesteKarte = SucheBesteKarteOberMitspieler
            Exit Function
        End If
        '---------------------------------------------------------------------------------------------------
        If Spielart = ErsterUndLetzterstich Then
            'Hier darf der Computer auch hohe Karten ausspielen, wenn es nicht der ErsterUndLetzterStich ist
            SucheBesteKarte = SucheBesteKarteErsterUndLetzter
            Exit Function
        End If
        '---------------------------------------------------------------------------------------------------
        If Spielart = RoterK�nig Then
            'Wenn abgeworfen werden kann, zuerst RoterK�nig abwerfen,
            'wenn zugegeben werden muss und es ist eine h�here Karte gespielt als RoterK�nig, dann RoterK�nig zugeben
            'wenn zugegeben werden muss und es ist eine niedrigere Karte gespielt als RoterK�nig, RoterK�nig nur wenn er der letzte von Rot ist
            SucheBesteKarte = SucheBesteKarteRoterK�nigMitspieler
            Exit Function
        End If
        '---------------------------------------------------------------------------------------------------
        If Spielart = Lorum Then
            'Lorum hat eigene Regeln
            
        End If
        '---------------------------------------------------------------------------------------------------
        'hier ist Spielart = Rot oder Stiche
        'der computer spielt nicht aus sondern muss zugeben oder abwerfen
        'Bei den ersten zwei Stichen darf nicht Rot abgeworfen werden und es muss die niedrigste zugegeben werden
        SucheBesteKarte = SucheBesteKarteRotOderSticheMitspieler
        Exit Function
    End If
End Function


Private Function SucheBesteKarteRotOderSticheAusspieler()
    'Links Mitte Rechts ist Ausspieler
    'finde die beste zum Ausspielen
    'Hier ist es Spielart Rot erster und zweiter Stich oder Stiche
    'Durchsuche alle Karten au�er rote bei den ersten zwei Stichen und finde die niedrigste zum Ausspielen
    'Gerbing 11.07.2017 Intelligenz-Verbesserung bei Spielart=Rot bei den ersten zwei Stichen die h�chste Nicht-Rot ausspielen

    Dim I As Integer
    Dim n As Integer
    Dim NiedrigsterWert As Integer
    Dim NiedrigstePosition As Integer
    Dim H�chsterWert As Integer
    Dim H�chstePosition As Integer

    NiedrigsterWert = 8                                             '8=Ass=Startwert
    If PlayerTurn = Links Then
        If Spielart <> Rot Then
            'es ist Spielart=Stiche
            For I = 1 To UBound(LinksCards)
                If LinksCards(I).cType <> 2 Then                        '2=Rot
                    If LinksCards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = LinksCards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteRotOderSticheAusspieler = NiedrigstePosition
            Exit Function
        Else
            'es ist Spielart=Rot
            H�chsterWert = 0                                            'Gerbing 11.07.2017
            n = 0
            For I = 1 To UBound(LinksCards)
                If LinksCards(I).cType <> 2 Then                        '2=Rot
                    'es ist keine Rot-Karte
                    If LinksCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = LinksCards(I).cValue
                        H�chstePosition = I
                    End If
                Else
                    'Vielleicht hat Links nur Rot                       'Gerbing 12.07.2017
                    n = n + 1
                    If n >= 7 Then
                        gstrMsg = "Links kann nur Rot ausspielen"
                        frmMsgBox.Show 1
                        SucheBesteKarteRotOderSticheAusspieler = I
                        Exit Function
                    End If
                End If
            Next I
            SucheBesteKarteRotOderSticheAusspieler = H�chstePosition
            Exit Function
        End If
    End If
    '----------------------------
    If PlayerTurn = Mitte Then
        If Spielart <> Rot Then
            'es ist Spielart=Stiche
            For I = 1 To UBound(MitteCards)
                If MitteCards(I).cType <> 2 Then                        '2=Rot
                    If MitteCards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = MitteCards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteRotOderSticheAusspieler = NiedrigstePosition
            Exit Function
        Else
            'es ist Spielart=Rot
            H�chsterWert = 0                                            'Gerbing 11.07.2017
            n = 0
            For I = 1 To UBound(MitteCards)
                If MitteCards(I).cType <> 2 Then                        '2=Rot
                    'es ist keine Rot-Karte
                    If MitteCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = MitteCards(I).cValue
                        H�chstePosition = I
                    End If
                Else
                    'Vielleicht hat mitte nur Rot                       'Gerbing 12.07.2017
                    n = n + 1
                    If n >= 7 Then
                        gstrMsg = "Mitte kann nur Rot ausspielen"
                        frmMsgBox.Show 1
                        SucheBesteKarteRotOderSticheAusspieler = I
                        Exit Function
                    End If
                End If
            Next I
            SucheBesteKarteRotOderSticheAusspieler = H�chstePosition
            Exit Function
        End If
    End If
    '----------------------------
    If PlayerTurn = Rechts Then
        If Spielart <> Rot Then
            'es ist Spielart=Stiche
            For I = 1 To UBound(RechtsCards)
                If RechtsCards(I).cType <> 2 Then                        '2=Rot
                    If RechtsCards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = RechtsCards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteRotOderSticheAusspieler = NiedrigstePosition
            Exit Function
        Else
            'es ist Spielart=Rot
            H�chsterWert = 0                                            'Gerbing 11.07.2017
            n = 0
            For I = 1 To UBound(RechtsCards)
                If RechtsCards(I).cType <> 2 Then                        '2=Rot
                    'es ist keine Rot-Karte
                    If RechtsCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = RechtsCards(I).cValue
                        H�chstePosition = I
                    End If
                Else
                    'Vielleicht hat rechts nur Rot                       'Gerbing 12.07.2017
                    n = n + 1
                    If n >= 7 Then
                        gstrMsg = "Rechts kann nur Rot ausspielen"
                        frmMsgBox.Show 1
                        SucheBesteKarteRotOderSticheAusspieler = I
                        Exit Function
                    End If
                End If
            Next I
            SucheBesteKarteRotOderSticheAusspieler = H�chstePosition
            Exit Function
        End If
    End If
End Function

Private Function SucheBesteKarteRotOderSticheAusspieler_N2()                                'N2=NachDemZweitenStich
    'Der Computer ist Ausspieler
    'finde die Position der besten Karte zum Ausspielen
    'Hier ist es Spielart (Rot nach dem zweiten Stich) oder Stiche oder ErsterUndLetzter
    'Finde Rot9 oder Rot8 oder Rot7                                                        'Gerbing 04.11.2017
    'wenn nicht gefunden, dann finde die niedrigste zum Ausspielen
    'Es ist sinnlos eine Farbe auszuspielen, die nur nur noch der Computer (Links Mitte Rechts) selbst hat
    '       dazu dient NachbehandlungAusspieler

    Dim I As Integer
    Dim NiedrigsterWert As Integer
    Dim NiedrigstePosition As Integer

    NiedrigsterWert = 8                                                         '8=Ass=Startwert
    NiedrigstePosition = 0
    If PlayerTurn = Links Then
        'Zuerst nach Rot9 oder Rot8 oder Rot7 suchen                            'Gerbing 04.11.2017
        For I = 1 To UBound(LinksCards)
            If LinksCards(I).cValue <= 3 And LinksCards(I).cType = 2 Then       'ist es Rot9
                NiedrigstePosition = I
                Exit For
            ElseIf LinksCards(I).cValue <= 2 And LinksCards(I).cType = 2 Then   'ist es Rot8
                NiedrigstePosition = I
                Exit For
            ElseIf LinksCards(I).cValue <= 1 And LinksCards(I).cType = 2 Then   'ist es Rot7
                NiedrigstePosition = I
                Exit For
            End If
        Next I
        If NiedrigstePosition <> 0 Then
            SucheBesteKarteRotOderSticheAusspieler_N2 = NiedrigstePosition
            Exit Function
        End If
        'dann finde die niedrigste zum Ausspielen
        For I = 1 To UBound(LinksCards)
            If LinksCards(I).cValue <= NiedrigsterWert Then
                NiedrigsterWert = LinksCards(I).cValue
                NiedrigstePosition = I
            End If
        Next I
        SucheBesteKarteRotOderSticheAusspieler_N2 = NachbehandlungAusspieler(LinksCards, NiedrigstePosition)
        Exit Function
    End If
    If PlayerTurn = Mitte Then
        'Zuerst nach Rot9 oder Rot8 oder Rot7 suchen                            'Gerbing 04.11.2017
        For I = 1 To UBound(MitteCards)
            If MitteCards(I).cValue <= 3 And MitteCards(I).cType = 2 Then       'ist es Rot9
                NiedrigstePosition = I
                Exit For
            ElseIf MitteCards(I).cValue <= 2 And MitteCards(I).cType = 2 Then   'ist es Rot8
                NiedrigstePosition = I
                Exit For
            ElseIf MitteCards(I).cValue <= 1 And MitteCards(I).cType = 2 Then   'ist es Rot7
                NiedrigstePosition = I
                Exit For
            End If
        Next I
        If NiedrigstePosition <> 0 Then
            SucheBesteKarteRotOderSticheAusspieler_N2 = NiedrigstePosition
            Exit Function
        End If
        'dann finde die niedrigste zum Ausspielen
        For I = 1 To UBound(MitteCards)
            If MitteCards(I).cValue <= NiedrigsterWert Then
                NiedrigsterWert = MitteCards(I).cValue
                NiedrigstePosition = I
            End If
        Next I
        SucheBesteKarteRotOderSticheAusspieler_N2 = NachbehandlungAusspieler(MitteCards, NiedrigstePosition)
        Exit Function
    End If
    If PlayerTurn = Rechts Then
        'Zuerst nach Rot9 oder Rot8 oder Rot7 suchen                            'Gerbing 04.11.2017
        For I = 1 To UBound(RechtsCards)
            If RechtsCards(I).cValue <= 3 And RechtsCards(I).cType = 2 Then       'ist es Rot9
                NiedrigstePosition = I
                Exit For
            ElseIf RechtsCards(I).cValue <= 2 And RechtsCards(I).cType = 2 Then   'ist es Rot8
                NiedrigstePosition = I
                Exit For
            ElseIf RechtsCards(I).cValue <= 1 And RechtsCards(I).cType = 2 Then   'ist es Rot7
                NiedrigstePosition = I
                Exit For
            End If
        Next I
        If NiedrigstePosition <> 0 Then
            SucheBesteKarteRotOderSticheAusspieler_N2 = NiedrigstePosition
            Exit Function
        End If
        'dann finde die niedrigste zum Ausspielen
        For I = 1 To UBound(RechtsCards)
            If RechtsCards(I).cValue <= NiedrigsterWert Then
                NiedrigsterWert = RechtsCards(I).cValue
                NiedrigstePosition = I
            End If
        Next I
        SucheBesteKarteRotOderSticheAusspieler_N2 = NachbehandlungAusspieler(RechtsCards, NiedrigstePosition)
        Exit Function
    End If
End Function

Private Function SucheBesteKarteRotOderSticheMitspieler()
    'Der Computer ist Mitspieler und muss zugeben oder abwerfen
    'Bei den ersten zwei Stichen darf nicht Rot abgeworfen werden und es muss die h�chste abgeworfen werden
    'Bei den ersten zwei Stichen bei spielart = Rot kann die h�chste Karte zugegeben werden Gerbing 09.07.2017
    'Ab 3.Stich ist vorrangig Rot abzuwerfen Gerbing 13.10.2017
    'Wenn eine Farbe zum erstenmal gespielt wird, ist zu pr�fen, ob der Computer von einer Farbe eine hohe und eine niedrige Karte hat.
    'Es ist in diesem Fall besser zuerst die hohe Karte zuzugeben
    '       dazu dient NachbehandlungMitspieler
    'Anstelle von DieH�chsteAbwerfen ist es manchmal besser DieEinzigeAbwerfen(wenn es nicht gerade die Sieben ist)
    '       dazu dient NachbehandlungAbwerfen
    'Wenn ich der 4.Zugeber(der letzte eines Stiches) bin, muss ich versuchen
    'geradeso unterhalb der h�chsten Karte zu bleiben, die auf dem Tisch liegt, nicht einfach die niedrigste
    '       dazu dient VorbehandlungRotOderStiche

    Dim I As Integer
    Dim NiedrigsterWert As Integer
    Dim NiedrigstePosition As Integer
    Dim H�chsterWert As Integer
    Dim H�chstePosition As Integer
    Dim AKType As Integer                                                   'AusgespielteKarteType
    Dim AKValue As Integer                                                   'AusgespielteKarteValue
    Dim blnFarbeVorhanden As Boolean
    
    blnFarbeVorhanden = False
    If Ausspieler = Links Then
        AKType = LinksGespielt.CardType
        AKValue = LinksGespielt.CardValue
    ElseIf Ausspieler = Mitte Then
        AKType = MitteGespielt.CardType
        AKValue = MitteGespielt.CardValue
    ElseIf Ausspieler = Rechts Then
        AKType = RechtsGespielt.CardType
        AKValue = RechtsGespielt.CardValue
    ElseIf Ausspieler = Spieler Then
        AKType = SpielerGespielt.CardType
        AKValue = SpielerGespielt.CardValue
    End If
'--------------------------------------------------------------------------------------------------
    If PlayerTurn = Links Then
        For I = 1 To UBound(LinksCards)
            If LinksCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            If AKIndex < 8 Then                                             'Die ersten zwei Stiche
                'Computer hat nicht die geforderte Farbe aber darf kein Rot abwerfen
                'Suche die h�chste aber nicht Rot
                H�chsterWert = 0
                For I = 1 To UBound(LinksCards)
                    If LinksCards(I).cType <> 2 Then                        '2=Rot
                        If LinksCards(I).cValue >= H�chsterWert Then
                            H�chsterWert = LinksCards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                SucheBesteKarteRotOderSticheMitspieler = NachbehandlungAbwerfen(LinksCards, H�chstePosition)
                Exit Function
            Else
                'es ist der dritte oder weitere Stich
                'Bei Rot vorrangig h�chste Rot abwerfen                     'Gerbing 13.10.2017
                If Spielart = Rot Then
                    H�chsterWert = 0
                    For I = 1 To UBound(LinksCards)
                        If LinksCards(I).cValue >= H�chsterWert And LinksCards(I).cType = 2 Then '2=Rot
                            H�chsterWert = LinksCards(I).cValue
                            H�chstePosition = I
                        End If
                    Next I
                    If H�chsterWert <> 0 Then
                        'Rot gefunden
                        SucheBesteKarteRotOderSticheMitspieler = H�chstePosition
                        Exit Function
                    End If
                End If
                'sonst h�chste Karte abwerfen
                H�chsterWert = 0
                For I = 1 To UBound(LinksCards)
                    If LinksCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = LinksCards(I).cValue
                        H�chstePosition = I
                    End If
                Next I
                SucheBesteKarteRotOderSticheMitspieler = NachbehandlungAbwerfen(LinksCards, H�chstePosition)
                Exit Function
            End If
        Else
            'Farbe ist vorhanden, es muss zugegeben werden
            If AKIndex < 8 Then
                'Bei den ersten zwei Stichen bei spielart = Rot kann die h�chste Karte zugegeben werden 'Gerbing 09.07.2017
                If Spielart = Rot Then
                    H�chsterWert = 0
                    For I = 1 To UBound(LinksCards)
                        If AKType = LinksCards(I).cType Then
                            If LinksCards(I).cValue >= H�chsterWert Then
                                H�chsterWert = LinksCards(I).cValue
                                H�chstePosition = I
                            End If
                        End If
                    Next I
                    SucheBesteKarteRotOderSticheMitspieler = H�chstePosition
                    Exit Function
                End If
            End If
            'Nach dem 2.Stich Rot oder ab 1.Stich Stiche, Farbe ist vorhanden                           'Gerbing 16.10.2017
            'Wenn ich der 4.Zugeber bin -> VorbehandlungRotOderStiche
            If blnMitteWarDran = True And blnRechtsWarDran = True And blnSpielerWarDran = True Then
                SucheBesteKarteRotOderSticheMitspieler = VorbehandlungRotOderStiche(LinksCards, PlayerTurn, AKType)
                Exit Function
            End If                                                                                      'Gerbing 16.10.2017
            'sonst die niedrigste zugeben
            NiedrigsterWert = 8                                             '8=Ass=Startwert
            For I = 1 To UBound(LinksCards)
                If AKType = LinksCards(I).cType Then
                    If LinksCards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = LinksCards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteRotOderSticheMitspieler = NachbehandlungMitspieler(LinksCards, NiedrigstePosition)
            Exit Function
        End If
    End If
'--------------------------------------------------------------------------------------------------
    If PlayerTurn = Mitte Then
        For I = 1 To UBound(MitteCards)
            If MitteCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            If AKIndex < 8 Then                                             'Die ersten zwei Stiche
                'Computer hat nicht die geforderte Farbe aber darf kein Rot abwerfen
                'Suche die h�chste aber nicht Rot
                H�chsterWert = 0
                For I = 1 To UBound(MitteCards)
                    If MitteCards(I).cType <> 2 Then                        '2=Rot
                        If MitteCards(I).cValue >= H�chsterWert Then
                            H�chsterWert = MitteCards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                SucheBesteKarteRotOderSticheMitspieler = NachbehandlungAbwerfen(MitteCards, H�chstePosition)
                Exit Function
            Else
                'es ist der dritte oder weitere Stich
                'Bei Rot vorrangig h�chste Rot abwerfen                     'Gerbing 13.10.2017
                If Spielart = Rot Then
                    H�chsterWert = 0
                    For I = 1 To UBound(MitteCards)
                        If MitteCards(I).cValue >= H�chsterWert And MitteCards(I).cType = 2 Then '2=Rot
                            H�chsterWert = MitteCards(I).cValue
                            H�chstePosition = I
                        End If
                    Next I
                    If H�chsterWert <> 0 Then
                        'Rot gefunden
                        SucheBesteKarteRotOderSticheMitspieler = H�chstePosition
                        Exit Function
                    End If
                End If
                'sonst h�chste Karte abwerfen
                H�chsterWert = 0
                For I = 1 To UBound(MitteCards)
                    If MitteCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = MitteCards(I).cValue
                        H�chstePosition = I
                    End If
                Next I
                SucheBesteKarteRotOderSticheMitspieler = NachbehandlungAbwerfen(MitteCards, H�chstePosition)
                Exit Function
            End If
        Else
            'Farbe ist vorhanden, es muss zugegeben werden
            If AKIndex < 8 Then
                'Bei den ersten zwei Stichen bei spielart = Rot kann die h�chste Karte zugegeben werden 'Gerbing 09.07.2017
                If Spielart = Rot Then
                    H�chsterWert = 0
                    For I = 1 To UBound(MitteCards)
                        If AKType = MitteCards(I).cType Then
                            If MitteCards(I).cValue >= H�chsterWert Then
                                H�chsterWert = MitteCards(I).cValue
                                H�chstePosition = I
                            End If
                        End If
                    Next I
                    SucheBesteKarteRotOderSticheMitspieler = H�chstePosition
                    Exit Function
                End If
            End If
            'Nach dem 2.Stich Rot oder ab 1.Stich Stiche, Farbe ist vorhanden                           'Gerbing 16.10.2017
            'Wenn ich der 4. Zugeber bin -> VorbehandlungRotOderStiche
            If blnRechtsWarDran = True And blnSpielerWarDran = True And blnLinksWarDran = True Then
                SucheBesteKarteRotOderSticheMitspieler = VorbehandlungRotOderStiche(MitteCards, PlayerTurn, AKType)
                Exit Function
            End If                                                                                      'Gerbing 16.10.2017
            'sonst die niedrigste zugeben
            NiedrigsterWert = 8                                             '8=Ass=Startwert
            For I = 1 To UBound(MitteCards)
                If AKType = MitteCards(I).cType Then
                    If MitteCards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = MitteCards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteRotOderSticheMitspieler = NachbehandlungMitspieler(MitteCards, NiedrigstePosition)
            Exit Function
        End If
    End If
'--------------------------------------------------------------------------------------------------
    If PlayerTurn = Rechts Then
        For I = 1 To UBound(RechtsCards)
            If RechtsCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            If AKIndex < 8 Then                                            'Die ersten zwei Stiche
                'Computer hat nicht die geforderte Farbe aber darf kein Rot abwerfen
                'Suche die h�chste aber nicht Rot
                H�chsterWert = 0
                For I = 1 To UBound(RechtsCards)
                    If RechtsCards(I).cType <> 2 Then                        '2=Rot
                        If RechtsCards(I).cValue >= H�chsterWert Then
                            H�chsterWert = RechtsCards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                SucheBesteKarteRotOderSticheMitspieler = NachbehandlungAbwerfen(RechtsCards, H�chstePosition)
                Exit Function
            Else
                'es ist der dritte oder weitere Stich
                'Bei Rot vorrangig h�chste Rot abwerfen                     'Gerbing 13.10.2017
                If Spielart = Rot Then
                    H�chsterWert = 0
                    For I = 1 To UBound(RechtsCards)
                        If RechtsCards(I).cValue >= H�chsterWert And RechtsCards(I).cType = 2 Then '2=Rot
                            H�chsterWert = RechtsCards(I).cValue
                            H�chstePosition = I
                        End If
                    Next I
                    If H�chsterWert <> 0 Then
                        'Rot gefunden
                        SucheBesteKarteRotOderSticheMitspieler = H�chstePosition
                        Exit Function
                    End If
                End If
                'sonst h�chste Karte abwerfen
                H�chsterWert = 0
                For I = 1 To UBound(RechtsCards)
                    If RechtsCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = RechtsCards(I).cValue
                        H�chstePosition = I
                    End If
                Next I
                SucheBesteKarteRotOderSticheMitspieler = NachbehandlungAbwerfen(RechtsCards, H�chstePosition)
                Exit Function
            End If
        Else
            'Farbe ist vorhanden, es muss zugegeben werden
            If AKIndex < 8 Then
                'Bei den ersten zwei Stichen bei spielart = Rot kann die h�chste Karte zugegeben werden 'Gerbing 09.07.2017
                If Spielart = Rot Then
                    H�chsterWert = 0
                    For I = 1 To UBound(RechtsCards)
                        If AKType = RechtsCards(I).cType Then
                            If RechtsCards(I).cValue >= H�chsterWert Then
                                H�chsterWert = RechtsCards(I).cValue
                                H�chstePosition = I
                            End If
                        End If
                    Next I
                    SucheBesteKarteRotOderSticheMitspieler = H�chstePosition
                    Exit Function
                End If
            End If
            'Nach dem 2.Stich Rot oder ab 1.Stich Stiche, Farbe ist vorhanden                           'Gerbing 16.10.2017
            'Wenn ich der 4. Zugeber bin -> VorbehandlungRotOderStiche
            If blnSpielerWarDran = True And blnLinksWarDran = True And blnMitteWarDran = True Then
                SucheBesteKarteRotOderSticheMitspieler = VorbehandlungRotOderStiche(RechtsCards, PlayerTurn, AKType)
                Exit Function
            End If                                                                                      'Gerbing 16.10.2017
            'sonst die niedrigste zugeben
            NiedrigsterWert = 8                                             '8=Ass=Startwert
            For I = 1 To UBound(RechtsCards)
                If AKType = RechtsCards(I).cType Then
                    If RechtsCards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = RechtsCards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteRotOderSticheMitspieler = NachbehandlungMitspieler(RechtsCards, NiedrigstePosition)
            Exit Function
        End If
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call term_global                                            'Gerbing 18.11.2017
    End                                                         'Gerbing 23.10.2017
End Sub

Private Sub lblRoserl_Click(Index As Integer)
    If blnNeuesSpielBeginnt = True Then
        If Ausspieler = Links Then
            'strAusspieler = "Links"
            lblLinks.Font.Bold = True                           'Gerbing 15.08.2017
            lblLinks.ForeColor = FarbeAusspieler                           'schwarz
            lblSpieler.ForeColor = &HFFFFFF                     'wei�
            lblMitte.ForeColor = &HFFFFFF
            lblRechts.ForeColor = &HFFFFFF
            lblMitte.Font.Bold = False
            lblRechts.Font.Bold = False
            lblSpieler.Font.Bold = False
        End If
        If Ausspieler = Mitte Then
            'strAusspieler = "Mitte"
            lblMitte.Font.Bold = True                           'Gerbing 15.08.2017
            lblMitte.ForeColor = FarbeAusspieler                           'schwarz
            lblLinks.ForeColor = &HFFFFFF                       'wei�
            lblSpieler.ForeColor = &HFFFFFF
            lblRechts.ForeColor = &HFFFFFF
            lblLinks.Font.Bold = False
            lblRechts.Font.Bold = False
            lblSpieler.Font.Bold = False
        End If
        If Ausspieler = Rechts Then
            'strAusspieler = "Rechts"
            lblRechts.Font.Bold = True                          'Gerbing 15.08.2017
            lblRechts.ForeColor = FarbeAusspieler                          'schwarz
            lblLinks.ForeColor = &HFFFFFF                       'wei�
            lblMitte.ForeColor = &HFFFFFF
            lblSpieler.ForeColor = &HFFFFFF
            lblLinks.Font.Bold = False
            lblMitte.Font.Bold = False
            lblSpieler.Font.Bold = False
        End If
        If Ausspieler = Spieler Then
            'strAusspieler = "Spieler"
            lblSpieler.Font.Bold = True                         'Gerbing 15.08.2017
            lblSpieler.ForeColor = FarbeAusspieler                         'schwarz
            lblLinks.ForeColor = &HFFFFFF                       'wei�
            lblMitte.ForeColor = &HFFFFFF
            lblRechts.ForeColor = &HFFFFFF
            lblLinks.Font.Bold = False
            lblMitte.Font.Bold = False
            lblRechts.Font.Bold = False
        End If
        If Spielart = Lorum Then
            blnNeuesSpielBeginnt = False
            If PlayerTurnRoserl = Links Then                    'Gerbing 14.10.2017
                Call cdLinks_Click(0)
            ElseIf PlayerTurnRoserl = Mitte Then                'Gerbing 14.10.2017
                Call cdMitte_Click(0)
            ElseIf PlayerTurnRoserl = Rechts Then               'Gerbing 14.10.2017
                Call cdRechts_Click(0)
            End If
            Exit Sub
        End If
        blnNeuesSpielBeginnt = False
        'Neues Spiel beginnt
        If AusspielerBeiSpielbeginn = Links Then
            'Links Mitte und Rechts m�ssen noch ausspielen
            Call LinksPlay
            Call MittePlay
            Call RechtsPlay
        End If
        If AusspielerBeiSpielbeginn = Mitte Then
            'Mitte und Rechts muss noch ausspielen
            Call MittePlay
            Call RechtsPlay
        End If
        If AusspielerBeiSpielbeginn = Rechts Then
            'Rechts muss noch ausspielen
            Call RechtsPlay
        End If
        If AusspielerBeiSpielbeginn = Spieler Then
            'Warten dass der Spieler ausspielt
            'MsgBox bringen wenn der Spieler 10 Sekunden schl�ft
            SchlafTimer.Enabled = True
        End If
    End If
End Sub

Private Function NachbehandlungAbwerfen(computercards() As cRoserl, H�chstePosition As Integer)
    'Anstelle von DieH�chsteAbwerfen ist es manchmal besser DieEinzigeAbwerfen
    '(wenn es nicht gerade die 7 oder 8 oder 9 ist) 'Gerbing 15.10.2017
    'dazu dient NachbehandlungAbwerfen
    'es ist sinnlos eine Farbe abzuwerfen, die niemand mehr hat(niemand mehr ausspielen kann)
    'au�er ich kann keine andere abwerfen, dazu dient der Aufruf von HatKeinerMehr
    
    'NachbehandlungAbwerfen soll herausfinden ob die Karte H�chstePosition lieber nicht abgeworfen werden soll
    'Wenn keine bessere Karte gefunden wird, dann doch die nehmen, die bei Aufruf der Funktion in H�chstePosition stand
    
    Dim SaveH�chstePosition As Integer
    Dim I As Integer
    Dim z�hlerMK As Integer                                             'z�hler Meine Karten
    Dim SRGEPosition As Integer                                         'Schell Rot Gr�n Eichel Position
    Dim rc As Integer

    SaveH�chstePosition = H�chstePosition
    'wenn die empfohlene Karte bereits die einzige ist, nichts tun
    z�hlerMK = 0
    For I = 1 To UBound(computercards)
        If computercards(I).cType = computercards(H�chstePosition).cType Then
            z�hlerMK = z�hlerMK + 1
        End If
    Next I
    If z�hlerMK = 1 Then
        'es ist die einzige, nichts tun
        NachbehandlungAbwerfen = SaveH�chstePosition
        Exit Function
    End If
    '------------------------------------------------------------------------------------------
    'z�hlerMK ist h�her als 1
    If computercards(H�chstePosition).cType = 1 Then
    'Die empfohlene Karte war Schell und sie war nicht die einzige, probiere es mit rot oder gr�n oder eichel, ob dies die einzige ist
        'aber nicht bei Spielart = Rot  erster oder zweiter Stich           'Gerbing 04.09.2017
        If AKIndex < 8 And Spielart <> Rot Then
            'mit Rot
            z�hlerMK = 0
            For I = 1 To UBound(computercards)
                If computercards(I).cType = 2 Then                          'cType=2=Rot
                    z�hlerMK = z�hlerMK + 1
                    SRGEPosition = I
                End If
            Next I
            If z�hlerMK = 1 Then
                'es gibt nur eine rot
                If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=
                    'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Rot
                    rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                    If rc <> 0 Then
                        'die hat keiner mehr
                        NachbehandlungAbwerfen = SaveH�chstePosition
                    Else
                        NachbehandlungAbwerfen = SRGEPosition
                    End If                                                  'Gerbing 15.10.2017
                    Exit Function
                End If
            End If
        End If
        '--------------------
        'mit Gr�n
        z�hlerMK = 0
        For I = 1 To UBound(computercards)
            If computercards(I).cType = 3 Then                              'cType=3=Gr�n
                z�hlerMK = z�hlerMK + 1
                SRGEPosition = I
            End If
        Next I
        If z�hlerMK = 1 Then
            'es gibt nur eine Gr�n
            If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Gr�n
                rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                If rc <> 0 Then
                    'die hat keiner mehr
                    NachbehandlungAbwerfen = SaveH�chstePosition
                Else
                    NachbehandlungAbwerfen = SRGEPosition
                End If                                                  'Gerbing 15.10.2017
                Exit Function
            End If
        End If
        '--------------------
        'mit Eichel
        z�hlerMK = 0
        For I = 1 To UBound(computercards)
            If computercards(I).cType = 4 Then                              'cType=4=Eichel
                z�hlerMK = z�hlerMK + 1
                SRGEPosition = I
            End If
        Next I
        If z�hlerMK = 1 Then
            'es gibt nur eine Eichel
            If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Eichel
                rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                If rc <> 0 Then
                    'die hat keiner mehr
                    NachbehandlungAbwerfen = SaveH�chstePosition
                Else
                    NachbehandlungAbwerfen = SRGEPosition
                End If                                                  'Gerbing 15.10.2017
                Exit Function
            End If
        End If
        'hier keine Einzelkarte gefunden
        'oder EinzelKarte gefunden, aber es ist die 7 oder 8 oder 9
        NachbehandlungAbwerfen = SaveH�chstePosition
        Exit Function
    End If
    '------------------------------------------------------------------------------------------
    If computercards(H�chstePosition).cType = 2 Then
    'Die empfohlene Karte war Rot und sie war nicht die einzige, probiere es mit schell oder gr�n oder eichel, ob dies die einzige ist
        'mit Schell
        z�hlerMK = 0
        For I = 1 To UBound(computercards)
            If computercards(I).cType = 1 Then                              'cType=1=Schell
                z�hlerMK = z�hlerMK + 1
                SRGEPosition = I
            End If
        Next I
        If z�hlerMK = 1 Then
            'es gibt nur eine schell
            If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Schell
                rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                If rc <> 0 Then
                    'die hat keiner mehr
                    NachbehandlungAbwerfen = SaveH�chstePosition
                Else
                    NachbehandlungAbwerfen = SRGEPosition
                End If                                                  'Gerbing 15.10.2017
                Exit Function
            End If
        End If
        '--------------------
        'mit Gr�n
        z�hlerMK = 0
        For I = 1 To UBound(computercards)
            If computercards(I).cType = 3 Then                              'cType=3=Gr�n
                z�hlerMK = z�hlerMK + 1
                SRGEPosition = I
            End If
        Next I
        If z�hlerMK = 1 Then
            'es gibt nur eine Gr�n
            If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Gr�n
                rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                If rc <> 0 Then
                    'die hat keiner mehr
                    NachbehandlungAbwerfen = SaveH�chstePosition
                Else
                    NachbehandlungAbwerfen = SRGEPosition
                End If                                                  'Gerbing 15.10.2017
                Exit Function
            End If
        End If
        '--------------------
        'mit Eichel
        z�hlerMK = 0
        For I = 1 To UBound(computercards)
            If computercards(I).cType = 4 Then                              'cType=4=Eichel
                z�hlerMK = z�hlerMK + 1
                SRGEPosition = I
            End If
        Next I
        If z�hlerMK = 1 Then
            'es gibt nur eine Eichel
            If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Eichel
                rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                If rc <> 0 Then
                    'die hat keiner mehr
                    NachbehandlungAbwerfen = SaveH�chstePosition
                Else
                    NachbehandlungAbwerfen = SRGEPosition
                End If                                                  'Gerbing 15.10.2017
                Exit Function
            End If
        End If
        'hier keine Einzelkarte gefunden
        'oder EinzelKarte gefunden, aber es ist die 7 oder 8 oder 9
        NachbehandlungAbwerfen = SaveH�chstePosition
        Exit Function
    End If
    '------------------------------------------------------------------------------------------
    If computercards(H�chstePosition).cType = 3 Then
    'Die empfohlene Karte war Gr�n und sie war nicht die einzige, probiere es mit schell oder Rot oder eichel, ob dies die einzige ist
        'mit Schell
        z�hlerMK = 0
        For I = 1 To UBound(computercards)
            If computercards(I).cType = 1 Then                              'cType=1=Schell
                z�hlerMK = z�hlerMK + 1
                SRGEPosition = I
            End If
        Next I
        If z�hlerMK = 1 Then
            'es gibt nur eine schell
            If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Schell
                rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                If rc <> 0 Then
                    'die hat keiner mehr
                    NachbehandlungAbwerfen = SaveH�chstePosition
                Else
                    NachbehandlungAbwerfen = SRGEPosition
                End If                                                  'Gerbing 15.10.2017
                Exit Function
            End If
        End If
        '--------------------
        'aber nicht bei Spielart = Rot  erster oder zweiter Stich           'Gerbing 04.09.2017
        If AKIndex < 8 And Spielart <> Rot Then
            'mit Rot
            z�hlerMK = 0
            For I = 1 To UBound(computercards)
                If computercards(I).cType = 2 Then                          'cType=2=rot
                    z�hlerMK = z�hlerMK + 1
                    SRGEPosition = I
                End If
            Next I
            If z�hlerMK = 1 Then
                'es gibt nur eine Rot
                If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                    'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Rot
                    rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                    If rc <> 0 Then
                        'die hat keiner mehr
                        NachbehandlungAbwerfen = SaveH�chstePosition
                    Else
                        NachbehandlungAbwerfen = SRGEPosition
                    End If                                                  'Gerbing 15.10.2017
                    Exit Function
                End If
            End If
        End If
        '--------------------
        'mit Eichel
        z�hlerMK = 0
        For I = 1 To UBound(computercards)
            If computercards(I).cType = 4 Then                              'cType=4=Eichel
                z�hlerMK = z�hlerMK + 1
                SRGEPosition = I
            End If
        Next I
        If z�hlerMK = 1 Then
            'es gibt nur eine Eichel
            If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Eichel
                rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                If rc <> 0 Then
                    'die hat keiner mehr
                    NachbehandlungAbwerfen = SaveH�chstePosition
                Else
                    NachbehandlungAbwerfen = SRGEPosition
                End If                                                  'Gerbing 15.10.2017
                Exit Function
            End If
        End If
        'hier keine Einzelkarte gefunden
        'oder EinzelKarte gefunden, aber es ist die 7 oder 8 oder 9
        NachbehandlungAbwerfen = SaveH�chstePosition
        Exit Function
    End If
    '------------------------------------------------------------------------------------------
    If computercards(H�chstePosition).cType = 4 Then
    'Die empfohlene Karte war Eichel und sie war nicht die einzige, probiere es mit schell oder Rot oder gr�n, ob dies die einzige ist
        'mit Schell
        z�hlerMK = 0
        For I = 1 To UBound(computercards)
            If computercards(I).cType = 1 Then                              'cType=1=Schell
                z�hlerMK = z�hlerMK + 1
                SRGEPosition = I
            End If
        Next I
        If z�hlerMK = 1 Then
            'es gibt nur eine schell
            If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Schell
                rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                If rc <> 0 Then
                    'die hat keiner mehr
                    NachbehandlungAbwerfen = SaveH�chstePosition
                Else
                    NachbehandlungAbwerfen = SRGEPosition
                End If                                                  'Gerbing 15.10.2017
                Exit Function
            End If
        End If
        '--------------------
        'aber nicht bei Spielart = Rot  erster oder zweiter Stich           'Gerbing 04.09.2017
        If AKIndex < 8 And Spielart <> Rot Then
            'mit Rot
            z�hlerMK = 0
            For I = 1 To UBound(computercards)
                If computercards(I).cType = 2 Then                          'cType=2=rot
                    z�hlerMK = z�hlerMK + 1
                    SRGEPosition = I
                End If
            Next I
            If z�hlerMK = 1 Then
                'es gibt nur eine Rot
                If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                    'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Rot
                    rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                    If rc <> 0 Then
                        'die hat keiner mehr
                        NachbehandlungAbwerfen = SaveH�chstePosition
                    Else
                        NachbehandlungAbwerfen = SRGEPosition
                    End If                                                  'Gerbing 15.10.2017
                    Exit Function
                End If
            End If
        End If
        '--------------------
        'mit Gr�n
        z�hlerMK = 0
        For I = 1 To UBound(computercards)
            If computercards(I).cType = 3 Then                              'cType=3=Gr�n
                z�hlerMK = z�hlerMK + 1
                SRGEPosition = I
            End If
        Next I
        If z�hlerMK = 1 Then
            'es gibt nur eine gr�n
            If computercards(SRGEPosition).cValue > 3 Then              'cValue=3=Neun
                'wenn es nicht gerade die 7 oder 8 oder 9 ist, nimm diese Gr�n
                rc = HatKeinerMehr(computercards, SRGEPosition)         'Gerbing 15.10.2017
                If rc <> 0 Then
                    'die hat keiner mehr
                    NachbehandlungAbwerfen = SaveH�chstePosition
                Else
                    NachbehandlungAbwerfen = SRGEPosition
                End If                                                  'Gerbing 15.10.2017
                Exit Function
            End If
        End If
    End If
    'hier keine Einzelkarte gefunden
    'oder EinzelKarte gefunden, aber es ist die 7
    NachbehandlungAbwerfen = SaveH�chstePosition
End Function

Private Function NachbehandlungAusspieler(computercards() As cRoserl, NiedrigstePosition As Integer)
    'wird aufgerufen von: Rot_N2 Ober Stiche ErsterUndLetzter RoterK�nig
    'Kontrolliere ob eine Farbe gespielt werden soll, von der bereits alle Karten ausgespielt sind
    'NachbehandlungAusspieler soll herausfinden ob die Karte NiedrigstePosition lieber nicht gespielt werden soll,
    'weil kein anderer sie hat
    'daf�r wird sie gekennzeichnet in cGespieltVon mit 1 als BeiErneutemSuchen�berspringen.
    'Wenn keine bessere Karte gefunden wird, dann doch die nehmen, die bei Aufruf der Funktion in NiedrigstePosition stand
    
    'Es ist sinnlos eine Farbe auszuspielen, die nur nur noch der Computer (Links Mitte Rechts) selbst hat
    
    Dim SaveNiedrigstePosition As Integer
    Dim I As Integer
    Dim z�hlerAK As Integer                                             'z�hler Ausgespielte Karten
    Dim z�hlerMK As Integer                                             'z�hler Meine Karten
    Dim NiedrigsterWert As Integer
    
    SaveNiedrigstePosition = NiedrigstePosition
    computercards(NiedrigstePosition).cGespieltVon = 1
    'Kontrolliere ob eine Farbe gespielt werden soll, von der bereits alle Karten ausgespielt sind
    'das ist dann der Fall wenn z�hlerAK plus z�hlerMK = 8
    z�hlerAK = 0
    For I = 1 To AKIndex
        If computercards(NiedrigstePosition).cType = AusgespielteKarten(I).cType Then
            z�hlerAK = z�hlerAK + 1
        End If
    Next I
    z�hlerMK = 0
    For I = 1 To UBound(computercards)
        If computercards(I).cType = computercards(NiedrigstePosition).cType Then
            z�hlerMK = z�hlerMK + 1
        End If
    Next I
    If z�hlerAK + z�hlerMK = 8 Then
        'weil schon alle ausgespielt eine andere Farbe suchen
        NiedrigstePosition = 0
        NiedrigsterWert = 8                                             '8=Ass=Startwert
        For I = 1 To UBound(computercards)
            If computercards(I).cGespieltVon <> 1 Then                  'Gerbing 13.10.2017 beim erneuten Suchen �berspringen
                'Gerbing 13.10.2017 und diese Farbe �berspringen die hat keiner mehr
                If computercards(SaveNiedrigstePosition).cType <> computercards(I).cType Then
                    If computercards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = computercards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            End If
        Next I
        If NiedrigstePosition = 0 Then
            'keine andere Farbe gefunden
            NachbehandlungAusspieler = SaveNiedrigstePosition
        Else
            NachbehandlungAusspieler = NiedrigstePosition
        End If
    Else
        'hier brauche ich keine andere Farbe zu suchen, es gibt noch welche von dieser Farbe
        NachbehandlungAusspieler = SaveNiedrigstePosition
    End If
End Function

Private Function NachbehandlungMitspieler(computercards() As cRoserl, NiedrigstePosition As Integer)
    'Wenn der Computer zugeben  muss und
    'wenn eine Farbe zum erstenmal gespielt wird, ist zu pr�fen, ob der Computer von dieser Farbe Farbe eine hohe und eine niedrige Karte hat.
    'und wenn die vorgeschlagene niedrige Karte 7 oder 8 oder 9 ist
    'Es ist in diesem Fall besser zuerst die hohe Karte zuzugeben
    'dazu dient NachbehandlungMitspieler
    
    'NachbehandlungMitspieler soll herausfinden ob die vorgeschlagene Karte NiedrigstePosition lieber nicht gespielt werden soll
    'Wenn keine bessere Karte gefunden wird, dann doch die nehmen, die bei Aufruf der Funktion in NiedrigstePosition stand
    
    Dim SaveNiedrigstePosition As Integer
    Dim I As Integer
    Dim z�hlerMK As Integer                                             'z�hler Meine Karten
    Dim H�chsterWert As Integer
    Dim H�chstePosition As Integer
    
    SaveNiedrigstePosition = NiedrigstePosition
    If computercards(NiedrigstePosition).cType = 1 Then                 '1=Schell
        'ich muss Schell zugeben
        If computercards(NiedrigstePosition).cValue = 1 Or computercards(NiedrigstePosition).cValue = 2 Or computercards(NiedrigstePosition).cValue = 3 Then
            ' es ist 7 oder 8 oder 9
        Else
            'es ist h�her als 9
            'ich nehme die vorgeschlagenen Karte
            NachbehandlungMitspieler = SaveNiedrigstePosition
            Exit Function
        End If
        If blnSchellZEM = True Then
            'es ist 7 oder 8 oder 9 und Schell zum ersten mal gespielt
            blnSchellZEM = False
            z�hlerMK = 0
            For I = 1 To UBound(computercards)
                If computercards(I).cType = 1 Then
                    z�hlerMK = z�hlerMK + 1
                End If
            Next I
            If z�hlerMK = 1 Then
                'ich habe nur eine Schell
                NachbehandlungMitspieler = SaveNiedrigstePosition
                Exit Function
            Else
                'ich habe mehrere Schell, suche die h�chste
                H�chsterWert = 0
                For I = 1 To UBound(computercards)
                    If computercards(I).cType = 1 Then                  '1=Schell
                        If computercards(I).cValue >= H�chsterWert Then
                            H�chsterWert = computercards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                NachbehandlungMitspieler = H�chstePosition
                Exit Function
            End If
        Else
            'es ist nicht zum ersten mal gespielt, nimm die vorgeschlagene
            NachbehandlungMitspieler = SaveNiedrigstePosition
            Exit Function
        End If
    End If
    '-------------------------------------------------------
    If computercards(NiedrigstePosition).cType = 2 Then                 '2=rot
        'ich muss Rot zugeben
        If computercards(NiedrigstePosition).cValue = 1 Or computercards(NiedrigstePosition).cValue = 2 Or computercards(NiedrigstePosition).cValue = 3 Then
            ' es ist 7 oder 8 oder 9
        Else
            'es ist h�her als 9
            'ich nehme die vorgeschlagenen Karte
            NachbehandlungMitspieler = SaveNiedrigstePosition
            Exit Function
        End If
        If blnRotZEM = True Then
            'es ist 7 oder 8 oder 9 und Rot zum ersten mal gespielt
            blnRotZEM = False
            z�hlerMK = 0
            For I = 1 To UBound(computercards)
                If computercards(I).cType = 1 Then
                    z�hlerMK = z�hlerMK + 1
                End If
            Next I
            If z�hlerMK = 1 Then
                'ich habe nur eine Rot
                NachbehandlungMitspieler = SaveNiedrigstePosition
                Exit Function
            Else
                'ich habe mehrere Rot, suche die h�chste
                H�chsterWert = 0
                For I = 1 To UBound(computercards)
                    If computercards(I).cType = 2 Then                  '2=Rot
                        If computercards(I).cValue >= H�chsterWert Then
                            H�chsterWert = computercards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                NachbehandlungMitspieler = H�chstePosition
                Exit Function
            End If
        Else
            'es ist nicht zum ersten mal gespielt, nimm die vorgeschlagene
            NachbehandlungMitspieler = SaveNiedrigstePosition
            Exit Function
        End If
    End If
    '-------------------------------------------------------
    If computercards(NiedrigstePosition).cType = 3 Then                 '3=Gr�n
        'ich muss Gr�n zugeben
        If computercards(NiedrigstePosition).cValue = 1 Or computercards(NiedrigstePosition).cValue = 2 Or computercards(NiedrigstePosition).cValue = 3 Then
            ' es ist 7 oder 8 oder 9
        Else
            'es ist h�her als 9
            'ich nehme die vorgeschlagenen Karte
            NachbehandlungMitspieler = SaveNiedrigstePosition
            Exit Function
        End If
        If blnGr�nZEM = True Then
            'es ist 7 oder 8 oder 9 und Gr�n zum ersten mal gespielt
            blnGr�nZEM = False
            z�hlerMK = 0
            For I = 1 To UBound(computercards)
                If computercards(I).cType = 1 Then
                    z�hlerMK = z�hlerMK + 1
                End If
            Next I
            If z�hlerMK = 1 Then
                'ich habe nur eine Gr�n
                NachbehandlungMitspieler = SaveNiedrigstePosition
                Exit Function
            Else
                'ich habe mehrere Gr�n, suche die h�chste
                H�chsterWert = 0
                For I = 1 To UBound(computercards)
                    If computercards(I).cType = 3 Then                  '3=gr�n
                        If computercards(I).cValue >= H�chsterWert Then
                            H�chsterWert = computercards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                NachbehandlungMitspieler = H�chstePosition
                Exit Function
            End If
        Else
            'es ist nicht zum ersten mal gespielt, nimm die vorgeschlagene
            NachbehandlungMitspieler = SaveNiedrigstePosition
            Exit Function
        End If
    End If
    '-------------------------------------------------------
    If computercards(NiedrigstePosition).cType = 4 Then                 '4=Eichel
        'ich muss Eichel zugeben
        If computercards(NiedrigstePosition).cValue = 1 Or computercards(NiedrigstePosition).cValue = 2 Or computercards(NiedrigstePosition).cValue = 3 Then
            ' es ist 7 oder 8 oder 9
        Else
            'es ist h�her als 9
            'ich nehme die vorgeschlagenen Karte
            NachbehandlungMitspieler = SaveNiedrigstePosition
            Exit Function
        End If
        If blnEichelZEM = True Then
            'es ist 7 oder 8 oder 9 und Eichel zum ersten mal gespielt
            blnEichelZEM = False
            z�hlerMK = 0
            For I = 1 To UBound(computercards)
                If computercards(I).cType = 1 Then
                    z�hlerMK = z�hlerMK + 1
                End If
            Next I
            If z�hlerMK = 1 Then
                'ich habe nur eine Eichel
                NachbehandlungMitspieler = SaveNiedrigstePosition
                Exit Function
            Else
                'ich habe mehrere Eichel, suche die h�chste
                H�chsterWert = 0
                For I = 1 To UBound(computercards)
                    If computercards(I).cType = 4 Then                  '4=Eichel
                        If computercards(I).cValue >= H�chsterWert Then
                            H�chsterWert = computercards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                NachbehandlungMitspieler = H�chstePosition
                Exit Function
            End If
        Else
            'es ist nicht zum ersten mal gespielt, nimm die vorgeschlagene
            NachbehandlungMitspieler = SaveNiedrigstePosition
            Exit Function
        End If
    End If
End Function

Private Function SucheBesteKarteOberAusspieler()
    'Der Computer ist Ausspieler
    'wenn ausgespielt werden muss, einen Ober nur, wenn es nur noch Ober gibt
    'ansonsten die niedrigste
    'aber nicht, wenn diese Farbe keiner mehr hat                   'Gerbing 15.10.2017

    Dim I As Integer
    Dim NiedrigsterWert As Integer
    Dim NiedrigstePosition As Integer
    Dim SaveNiedrigstePosition As Integer

    NiedrigsterWert = 8                                             '8=Ass=Startwert
    NiedrigstePosition = 0
    If PlayerTurn = Links Then
        For I = 1 To UBound(LinksCards)
            If LinksCards(I).cValue <> 6 Then                        '6=Ober
                If LinksCards(I).cValue <= NiedrigsterWert Then
                    NiedrigsterWert = LinksCards(I).cValue
                    NiedrigstePosition = I
                End If
            End If
        Next I
        If NiedrigstePosition = 0 Then
            'hier gibt es nur Ober
            NiedrigstePosition = 1
        End If
        SaveNiedrigstePosition = NiedrigstePosition                  'Gerbing 15.10.2017
        SucheBesteKarteOberAusspieler = NachbehandlungAusspieler(LinksCards, NiedrigstePosition)    'Gerbing 15.10.2017
        If LinksCards(SucheBesteKarteOberAusspieler).cValue = 6 Then
            SucheBesteKarteOberAusspieler = SaveNiedrigstePosition  'Gerbing 15.10.2017
        End If
        Exit Function
    End If
    If PlayerTurn = Mitte Then
        For I = 1 To UBound(MitteCards)
            If MitteCards(I).cValue <> 6 Then                       '6=Ober
                If MitteCards(I).cValue <= NiedrigsterWert Then
                    NiedrigsterWert = MitteCards(I).cValue
                    NiedrigstePosition = I
                End If
            End If
        Next I
        If NiedrigstePosition = 0 Then
            'hier gibt es nur Ober
            NiedrigstePosition = 1
        End If
        SaveNiedrigstePosition = NiedrigstePosition                  'Gerbing 15.10.2017
        SucheBesteKarteOberAusspieler = NachbehandlungAusspieler(MitteCards, NiedrigstePosition)    'Gerbing 15.10.2017
        If MitteCards(SucheBesteKarteOberAusspieler).cValue = 6 Then
            SucheBesteKarteOberAusspieler = SaveNiedrigstePosition  'Gerbing 15.10.2017
        End If
        Exit Function
    End If
    If PlayerTurn = Rechts Then
        For I = 1 To UBound(RechtsCards)
            If RechtsCards(I).cValue <> 6 Then                      '6=Ober
                If RechtsCards(I).cValue <= NiedrigsterWert Then
                    NiedrigsterWert = RechtsCards(I).cValue
                    NiedrigstePosition = I
                End If
            End If
        Next I
        If NiedrigstePosition = 0 Then
            'hier gibt es nur Ober
            NiedrigstePosition = 1
        End If
        SaveNiedrigstePosition = NiedrigstePosition                  'Gerbing 15.10.2017
        SucheBesteKarteOberAusspieler = NachbehandlungAusspieler(RechtsCards, NiedrigstePosition)    'Gerbing 15.10.2017
        If RechtsCards(SucheBesteKarteOberAusspieler).cValue = 6 Then
            SucheBesteKarteOberAusspieler = SaveNiedrigstePosition  'Gerbing 15.10.2017
        End If
        Exit Function
    End If
End Function

Private Function SucheBesteKarteOberMitspieler()
    'Der Computer ist Mitspieler
    'Wenn abgeworfen werden kann, zuerst einen Ober abwerfen,
    'wenn zugegeben werden muss und es ist eine h�here Karte auf dem Tisch als Ober, dann den Ober zugeben
    'wenn zugegeben werden muss und es sind nur niedrigere Karte auf dem Tisch als Ober, den Ober nur wenn er der letzte dieser Farbe ist
    'Wenn der Ober noch nicht gespielt wurde und nicht auf dem Tisch liegt
    'und ich bin der 4.Zugeber(3 Karten liegen auf dem Tisch) -> VorbehandlungOber

    
    Dim AI As Integer
    Dim I As Integer
    Dim NiedrigsterWert As Integer
    Dim NiedrigstePosition As Integer
    Dim H�chsterWert As Integer
    Dim H�chstePosition As Integer
    Dim AKType As Integer                                                   'AusgespielteKarteType
    Dim AKValue As Integer                                                   'AusgespielteKarteValue
    Dim OberPosition As Integer
    Dim NichtOberPosition As Integer
    Dim blnFarbeVorhanden As Boolean
    Dim blnH�herAlsOber As Boolean
    Dim blnOberAufDemTisch As Boolean
    Dim blnOberSchonGespielt As Boolean
    
    blnFarbeVorhanden = False
    If Ausspieler = Links Then
        AKType = LinksGespielt.CardType
        AKValue = LinksGespielt.CardValue
    ElseIf Ausspieler = Mitte Then
        AKType = MitteGespielt.CardType
        AKValue = MitteGespielt.CardValue
    ElseIf Ausspieler = Rechts Then
        AKType = RechtsGespielt.CardType
        AKValue = RechtsGespielt.CardValue
    ElseIf Ausspieler = Spieler Then
        AKType = SpielerGespielt.CardType
        AKValue = SpielerGespielt.CardValue
    End If
    
    'liegt eine h�here Karte auf dem Tisch als Ober?                            'Gerbing 06.07.2017
    blnH�herAlsOber = False
    If (SpielerGespielt.Visible And SpielerGespielt.CardType = AKType And SpielerGespielt.CardValue > 6) Or _
        (LinksGespielt.Visible And LinksGespielt.CardType = AKType And LinksGespielt.CardValue > 6) Or _
        (MitteGespielt.Visible And MitteGespielt.CardType = AKType And MitteGespielt.CardValue > 6) Or _
        (RechtsGespielt.Visible And RechtsGespielt.CardType = AKType And RechtsGespielt.CardValue > 6) Then
        blnH�herAlsOber = True
    End If
    'liegt der Ober auf dem Tisch?                                              'Gerbing 06.07.2017
    blnOberAufDemTisch = False
    If (SpielerGespielt.Visible And SpielerGespielt.CardType = AKType And SpielerGespielt.CardValue = 6) Or _
        (LinksGespielt.Visible And LinksGespielt.CardType = AKType And LinksGespielt.CardValue = 6) Or _
        (MitteGespielt.Visible And MitteGespielt.CardType = AKType And MitteGespielt.CardValue = 6) Or _
        (RechtsGespielt.Visible And RechtsGespielt.CardType = AKType And RechtsGespielt.CardValue = 6) Then
        blnOberAufDemTisch = True
    End If
    'ist der Ober bereits in einem fr�heren Stich gespielt?
    blnOberSchonGespielt = False                                                'Gerbing 16.10.2017
    For I = 1 To AKIndex
        If AusgespielteKarten(I).cType = AKType Then
            If AusgespielteKarten(I).cValue = 6 Then                            '6=Ober
                blnOberSchonGespielt = True
            End If
        End If
    Next I
    If PlayerTurn = Links Then
        For I = 1 To UBound(LinksCards)
            If LinksCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            'Computer hat nicht die geforderte Farbe, soll vorrangig einen Ober abwerfen
            'Suche einen Ober, sonst die h�chste
            H�chsterWert = 0
            For I = 1 To UBound(LinksCards)
                If LinksCards(I).cValue = 6 Then                        '6=Ober
                    H�chstePosition = I
                    Exit For                                            'Gerbing 06.07.2017
                Else
                    If LinksCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = LinksCards(I).cValue
                        H�chstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteOberMitspieler = H�chstePosition
            Exit Function
        Else
            'es muss zugegeben werden,
            'Wenn der Ober noch nicht gespielt wurde und nicht auf dem Tisch liegt
            'und ich bin der 4.Zugeber(3 Karten liegen auf dem Tisch) -> VorbehandlungOber
            If blnOberSchonGespielt = False And blnOberAufDemTisch = False And blnMitteWarDran = True And blnRechtsWarDran = True And blnSpielerWarDran = True Then
                SucheBesteKarteOberMitspieler = VorbehandlungOber(LinksCards, PlayerTurn, AKType)   'Gerbing 16.10.2017
                If SucheBesteKarteOberMitspieler <> 0 Then              'Gerbing 28.11.2017
                    Exit Function
                End If
            End If
            'sonst, Farbe ist vorhanden, es muss zugegeben werden, ich bin nicht der 4.Zugeber
            'wenn zugegeben werden muss und es ist eine h�here Karte auf dem Tisch als Ober, dann den Ober zugeben, falls ich ihn habe
            NiedrigsterWert = 8                                         '8=Ass=Startwert
            OberPosition = 0
            If blnH�herAlsOber = True Then                                                  'Gerbing 06.07.2017
                'es ist eine h�here Karte auf dem Tisch als Ober
                For I = 1 To UBound(LinksCards)
                    If LinksCards(I).cValue = 6 And LinksCards(I).cType = AKType Then       '6=Ober
                        OberPosition = I
                        Exit For                                                            'Gerbing 06.07.2017
                    Else
                        If LinksCards(I).cValue <= NiedrigsterWert And LinksCards(I).cType = AKType Then 'Gerbing 08.07.2017
                            NiedrigsterWert = LinksCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                Next I
                If OberPosition <> 0 Then
                    'ich gebe den Ober zu
                    NiedrigstePosition = OberPosition
                End If
                SucheBesteKarteOberMitspieler = NiedrigstePosition
                Exit Function
            End If
            '-----------------------------------------------------
            'wenn zugegeben werden muss und es liegt der Ober auf dem Tisch, eine niedrige Karte spielen falls ich eine habe
            NiedrigsterWert = 8                                         '8=Ass=Startwert
            NiedrigstePosition = 0
            If blnOberAufDemTisch = True Then                                       '6=Ober Gerbing 06.07.2017
                For I = 1 To UBound(LinksCards)
                    If LinksCards(I).cType = AKType Then
                        If LinksCards(I).cValue <= NiedrigsterWert Then
                            NiedrigsterWert = LinksCards(I).cValue
                            NiedrigstePosition = I
                        Else
                            NichtOberPosition = I
                        End If
                    End If
                Next I
                If NiedrigstePosition <> 0 Then
                    SucheBesteKarteOberMitspieler = NiedrigstePosition
                    Exit Function
                Else
                    SucheBesteKarteOberMitspieler = NichtOberPosition
                    Exit Function
                End If
            End If
            '-----------------------------------------------------
            'wenn zugegeben werden muss und es ist wurde eine niedrigere Karte als Ober gespielt,
            'eine niedrige Karte spielen falls ich eine habe
            'sonst eine h�here als den Ober                              'Gerbing 06.07.2017
            NiedrigsterWert = 8                                         '8=Ass=Startwert
            NiedrigstePosition = 0
            For I = 1 To UBound(LinksCards)
                If LinksCards(I).cType = AKType Then
                    If LinksCards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = LinksCards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            Next I
            If LinksCards(NiedrigstePosition).cValue = 6 Then         '6=Ober Gerbing 06.07.2017
                'den Ober nicht spielen wenn noch eine h�here Karte da ist
                H�chsterWert = 6
                H�chstePosition = 0
                For I = 1 To UBound(LinksCards)
                    If LinksCards(I).cType = AKType Then
                        If LinksCards(I).cValue >= H�chsterWert Then
                            H�chsterWert = LinksCards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                If H�chstePosition <> 0 Then
                    NiedrigstePosition = H�chstePosition
                End If
            End If
            SucheBesteKarteOberMitspieler = NiedrigstePosition
            Exit Function
        End If
    End If
'--------------------------------------------------------------------------------------------------
    If PlayerTurn = Mitte Then
        For I = 1 To UBound(MitteCards)
            If MitteCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            'Computer hat nicht die geforderte Farbe, soll vorrangig einen Ober abwerfen
            'Suche einen Ober, sonst die h�chste
            H�chsterWert = 0
            For I = 1 To UBound(MitteCards)
                If MitteCards(I).cValue = 6 Then                       '6=Ober
                    H�chstePosition = I
                    Exit For                                            'Gerbing 06.07.2017
                Else
                    If MitteCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = MitteCards(I).cValue
                        H�chstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteOberMitspieler = H�chstePosition
            Exit Function
        Else
            'Wenn der Ober noch nicht gespielt wurde und ich bin der 4.Zugeber(3 Karten liegen auf dem Tisch) -> VorbehandlungOber
            If blnOberSchonGespielt = False And blnOberAufDemTisch = False And blnRechtsWarDran = True And blnSpielerWarDran = True And blnLinksWarDran = True Then
                SucheBesteKarteOberMitspieler = VorbehandlungOber(MitteCards, PlayerTurn, AKType) 'Gerbing 16.10.2017
                If SucheBesteKarteOberMitspieler <> 0 Then              'Gerbing 28.11.2017
                    Exit Function
                End If
            End If
            'sonst, Farbe ist vorhanden, es muss zugegeben werden, ich bin nicht der 4.Zugeber
            'wenn zugegeben werden muss und es ist eine h�here Karte auf dem Tisch als Ober, dann den Ober zugeben, falls ich ihn habe
            NiedrigsterWert = 8                                         '8=Ass=Startwert
            OberPosition = 0
            If blnH�herAlsOber = True Then                                                  'Gerbing 06.07.2017
                'es ist eine h�here Karte auf dem Tisch als Ober
                For I = 1 To UBound(MitteCards)
                    If MitteCards(I).cValue = 6 And MitteCards(I).cType = AKType Then       '6=Ober
                        OberPosition = I
                        Exit For                                                            'Gerbing 06.07.2017
                    Else
                        If MitteCards(I).cValue <= NiedrigsterWert And MitteCards(I).cType = AKType Then 'Gerbing 08.07.2017
                            NiedrigsterWert = MitteCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                Next I
                If OberPosition <> 0 Then
                    'ich gebe den Ober zu
                    NiedrigstePosition = OberPosition
                End If
                SucheBesteKarteOberMitspieler = NiedrigstePosition
                Exit Function
            End If
            '-----------------------------------------------------
            'wenn zugegeben werden muss und es liegt der Ober auf dem Tisch, eine niedrige Karte spielen falls ich eine habe
            NiedrigsterWert = 8                                         '8=Ass=Startwert
            NiedrigstePosition = 0
            If blnOberAufDemTisch = True Then                                       '6=Ober Gerbing 06.07.2017
                For I = 1 To UBound(MitteCards)
                    If MitteCards(I).cType = AKType Then
                        If MitteCards(I).cValue <= NiedrigsterWert Then
                            NiedrigsterWert = MitteCards(I).cValue
                            NiedrigstePosition = I
                        Else
                            NichtOberPosition = I
                        End If
                    End If
                Next I
                If NiedrigstePosition <> 0 Then
                    SucheBesteKarteOberMitspieler = NiedrigstePosition
                    Exit Function
                Else
                    SucheBesteKarteOberMitspieler = NichtOberPosition
                    Exit Function
                End If
            End If
            '-----------------------------------------------------
            'wenn zugegeben werden muss und es ist wurde eine niedrigere Karte als Ober gespielt,
            'eine niedrige Karte spielen falls ich eine habe
            'sonst eine h�here als den Ober                              'Gerbing 06.07.2017
            NiedrigsterWert = 8                                         '8=Ass=Startwert
            NiedrigstePosition = 0
            For I = 1 To UBound(MitteCards)
                If MitteCards(I).cType = AKType Then
                    If MitteCards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = MitteCards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            Next I
            If MitteCards(NiedrigstePosition).cValue = 6 Then         '6=Ober Gerbing 06.07.2017
                'den Ober nicht spielen wenn noch eine h�here Karte da ist
                H�chsterWert = 6
                H�chstePosition = 0
                For I = 1 To UBound(MitteCards)
                    If MitteCards(I).cType = AKType Then
                        If MitteCards(I).cValue >= H�chsterWert Then
                            H�chsterWert = MitteCards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                If H�chstePosition <> 0 Then
                    NiedrigstePosition = H�chstePosition
                End If
            End If
            SucheBesteKarteOberMitspieler = NiedrigstePosition
            Exit Function
        End If
    End If
'--------------------------------------------------------------------------------------------------
    If PlayerTurn = Rechts Then
        For I = 1 To UBound(RechtsCards)
            If RechtsCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            'Computer hat nicht die geforderte Farbe, soll vorrangig einen Ober abwerfen
            'Suche einen Ober, sonst die h�chste
            H�chsterWert = 0
            For I = 1 To UBound(RechtsCards)
                If RechtsCards(I).cValue = 6 Then                       '6=Ober
                    H�chstePosition = I
                    Exit For                                            'Gerbing 06.07.2017
                Else
                    If RechtsCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = RechtsCards(I).cValue
                        H�chstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteOberMitspieler = H�chstePosition
            Exit Function
        Else
            'Wenn der Ober noch nicht gespielt wurde und ich bin der 4.Zugeber(3 Karten liegen auf dem Tisch) -> VorbehandlungOber
            If blnOberSchonGespielt = False And blnOberAufDemTisch = False And blnSpielerWarDran = True And blnLinksWarDran = True And blnRechtsWarDran = True Then
                SucheBesteKarteOberMitspieler = VorbehandlungOber(RechtsCards, PlayerTurn, AKType) 'Gerbing 16.10.2017
                If SucheBesteKarteOberMitspieler <> 0 Then              'Gerbing 28.11.2017
                    Exit Function
                End If
            End If
            'sonst, Farbe ist vorhanden, es muss zugegeben werden, ich bin nicht der 4.Zugeber
            'wenn zugegeben werden muss und es ist eine h�here Karte auf dem Tisch als Ober, dann den Ober zugeben, falls ich ihn habe
            NiedrigsterWert = 8                                         '8=Ass=Startwert
            OberPosition = 0
            If blnH�herAlsOber = True Then                                                  'Gerbing 06.07.2017
                'es ist eine h�here Karte auf dem Tisch als Ober
                For I = 1 To UBound(RechtsCards)
                    If RechtsCards(I).cValue = 6 And RechtsCards(I).cType = AKType Then     '6=Ober
                        OberPosition = I
                        Exit For                                                            'Gerbing 06.07.2017
                    Else
                        If RechtsCards(I).cValue <= NiedrigsterWert And RechtsCards(I).cType = AKType Then 'Gerbing 08.07.2017
                            NiedrigsterWert = RechtsCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                Next I
                If OberPosition <> 0 Then
                    'ich gebe den Ober zu
                    NiedrigstePosition = OberPosition
                End If
                SucheBesteKarteOberMitspieler = NiedrigstePosition
                Exit Function
            End If
            '-----------------------------------------------------
            'wenn zugegeben werden muss und es liegt der Ober auf dem Tisch, eine niedrige Karte spielen falls ich eine habe
            NiedrigsterWert = 8                                         '8=Ass=Startwert
            NiedrigstePosition = 0
            If blnOberAufDemTisch = True Then                                       '6=Ober Gerbing 06.07.2017
                For I = 1 To UBound(RechtsCards)
                    If RechtsCards(I).cType = AKType Then
                        If RechtsCards(I).cValue <= NiedrigsterWert Then
                            NiedrigsterWert = RechtsCards(I).cValue
                            NiedrigstePosition = I
                        Else
                            NichtOberPosition = I
                        End If
                    End If
                Next I
                If NiedrigstePosition <> 0 Then
                    SucheBesteKarteOberMitspieler = NiedrigstePosition
                    Exit Function
                Else
                    SucheBesteKarteOberMitspieler = NichtOberPosition
                    Exit Function
                End If
            End If
            '-----------------------------------------------------
            'wenn zugegeben werden muss und es ist wurde eine niedrigere Karte als Ober gespielt,
            'eine niedrige Karte spielen falls ich eine habe
            'sonst eine h�here als den Ober                              'Gerbing 06.07.2017
            NiedrigsterWert = 8                                         '8=Ass=Startwert
            NiedrigstePosition = 0
            For I = 1 To UBound(RechtsCards)
                If RechtsCards(I).cType = AKType Then
                    If RechtsCards(I).cValue <= NiedrigsterWert Then
                        NiedrigsterWert = RechtsCards(I).cValue
                        NiedrigstePosition = I
                    End If
                End If
            Next I
            If RechtsCards(NiedrigstePosition).cValue = 6 Then         '6=Ober Gerbing 06.07.2017
                'den Ober nicht spielen wenn noch eine h�here Karte da ist
                H�chsterWert = 6
                H�chstePosition = 0
                For I = 1 To UBound(RechtsCards)
                    If RechtsCards(I).cType = AKType Then
                        If RechtsCards(I).cValue >= H�chsterWert Then
                            H�chsterWert = RechtsCards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                Next I
                If H�chstePosition <> 0 Then
                    NiedrigstePosition = H�chstePosition
                End If
            End If
            SucheBesteKarteOberMitspieler = NiedrigstePosition
            Exit Function
        End If
    End If
End Function

Private Function SucheBesteKarteRoterK�nigAusspieler()
    'Der Computer ist Ausspieler
    'wenn ausgespielt werden muss, RoterK�nig nur, wenn es die letzte Karte ist
    'ansonsten die niedrigste
    'aber nicht, wenn diese Farbe keiner mehr hat                   'Gerbing 15.10.2017
    
    Dim I As Integer
    Dim NiedrigsterWert As Integer
    Dim NiedrigstePosition As Integer
    Dim RoterK�nigPosition As Integer
    Dim SaveNiedrigstePosition As Integer

    NiedrigsterWert = 8                                             '8=Ass=Startwert
    NiedrigstePosition = 0
    If PlayerTurn = Links Then
        For I = 1 To UBound(LinksCards)
            If LinksCards(I).cValue = 7 And LinksCards(I).cType = 2 Then    'RoterK�nig
                RoterK�nigPosition = I
            Else
                If LinksCards(I).cValue <= NiedrigsterWert Then
                    NiedrigsterWert = LinksCards(I).cValue
                    NiedrigstePosition = I
                End If
            End If
        Next I
        If NiedrigstePosition = 0 Then
            'hier gibt es nur RoterK�nig
            NiedrigstePosition = 1
        End If
        SaveNiedrigstePosition = NiedrigstePosition                  'Gerbing 15.10.2017
        SucheBesteKarteRoterK�nigAusspieler = NachbehandlungAusspieler(LinksCards, NiedrigstePosition)    'Gerbing 15.10.2017
        If LinksCards(SucheBesteKarteOberAusspieler).cValue = 7 And LinksCards(SucheBesteKarteOberAusspieler).cType = 2 Then    'RoterK�nig
            SucheBesteKarteRoterK�nigAusspieler = SaveNiedrigstePosition  'Gerbing 15.10.2017
        End If
        Exit Function
    End If
    If PlayerTurn = Mitte Then
        For I = 1 To UBound(MitteCards)
            If MitteCards(I).cValue = 7 And MitteCards(I).cType = 2 Then    'RoterK�nig
                RoterK�nigPosition = I
            Else
                If MitteCards(I).cValue <= NiedrigsterWert Then
                    NiedrigsterWert = MitteCards(I).cValue
                    NiedrigstePosition = I
                End If
            End If
        Next I
        If NiedrigstePosition = 0 Then
            'hier gibt es nur RoterK�nig
            NiedrigstePosition = 1
        End If
        SaveNiedrigstePosition = NiedrigstePosition                  'Gerbing 15.10.2017
        SucheBesteKarteRoterK�nigAusspieler = NachbehandlungAusspieler(MitteCards, NiedrigstePosition)    'Gerbing 15.10.2017
        If MitteCards(SucheBesteKarteOberAusspieler).cValue = 7 And MitteCards(SucheBesteKarteOberAusspieler).cType = 2 Then    'RoterK�nig
            SucheBesteKarteRoterK�nigAusspieler = SaveNiedrigstePosition  'Gerbing 15.10.2017
        End If
        Exit Function
    End If
    If PlayerTurn = Rechts Then
        For I = 1 To UBound(RechtsCards)
            If RechtsCards(I).cValue = 7 And RechtsCards(I).cType = 2 Then    'RoterK�nig
                RoterK�nigPosition = I
            Else
                If RechtsCards(I).cValue <= NiedrigsterWert Then
                    NiedrigsterWert = RechtsCards(I).cValue
                    NiedrigstePosition = I
                End If
            End If
        Next I
        If NiedrigstePosition = 0 Then
            'hier gibt es nur RoterK�nig
            NiedrigstePosition = 1
        End If
        SaveNiedrigstePosition = NiedrigstePosition                  'Gerbing 15.10.2017
        SucheBesteKarteRoterK�nigAusspieler = NachbehandlungAusspieler(RechtsCards, NiedrigstePosition)    'Gerbing 15.10.2017
        If RechtsCards(SucheBesteKarteOberAusspieler).cValue = 7 And RechtsCards(SucheBesteKarteOberAusspieler).cType = 2 Then    'RoterK�nig
            SucheBesteKarteRoterK�nigAusspieler = SaveNiedrigstePosition  'Gerbing 15.10.2017
        End If
        Exit Function
    End If
End Function

Private Function SucheBesteKarteRoterK�nigMitspieler()
    'Der Computer ist Mitspieler
    'Wenn abgeworfen werden kann, zuerst RoterK�nig abwerfen,
    'wenn zugegeben werden muss und es ist eine h�here Karte gespielt als RoterK�nig, dann RoterK�nig zugeben
    'wenn zugegeben werden muss und es ist eine niedrigere Karte gespielt als RoterK�nig, RoterK�nig nur wenn er der letzte von Rot ist
    'Wenn der RoterK�nig noch nicht gespielt wurde und nicht auf dem Tisch liegt
    'und ich bin der 4.Zugeber(3 Karten liegen auf dem Tisch) -> VorbehandlungRoterK�nig(die h�chste Rot zugeben Ass)
    
    Dim AI As Integer
    Dim I As Integer
    Dim NiedrigsterWert As Integer
    Dim NiedrigstePosition As Integer
    Dim H�chsterWert As Integer
    Dim H�chstePosition As Integer
    Dim AKType As Integer                                                   'AusgespielteKarteType
    Dim AKValue As Integer                                                   'AusgespielteKarteValue
    Dim K�nigPosition As Integer
    Dim blnFarbeVorhanden As Boolean
    Dim blnRoterK�nigAufDemTisch As Boolean
    Dim blnRotAssAufDemTisch As Boolean                                     'Gerbing 08.12.2017
'    Dim blnRoterK�nigSchonGespielt As Boolean

    
    blnFarbeVorhanden = False
    If Ausspieler = Links Then
        AKType = LinksGespielt.CardType
        AKValue = LinksGespielt.CardValue
    ElseIf Ausspieler = Mitte Then
        AKType = MitteGespielt.CardType
        AKValue = MitteGespielt.CardValue
    ElseIf Ausspieler = Rechts Then
        AKType = RechtsGespielt.CardType
        AKValue = RechtsGespielt.CardValue
    ElseIf Ausspieler = Spieler Then
        AKType = SpielerGespielt.CardType
        AKValue = SpielerGespielt.CardValue
    End If
    'liegt RotAss auf dem Tisch?                                                    'Gerbing 08.12.2017 19.12.2017
    blnRotAssAufDemTisch = False
    If (SpielerGespielt.Visible And SpielerGespielt.CardType = 2 And SpielerGespielt.CardValue = 8) Or _
        (LinksGespielt.Visible And LinksGespielt.CardType = 2 And LinksGespielt.CardValue = 8) Or _
        (MitteGespielt.Visible And MitteGespielt.CardType = 2 And MitteGespielt.CardValue = 8) Or _
        (RechtsGespielt.Visible And RechtsGespielt.CardType = 2 And RechtsGespielt.CardValue = 8) Then
        blnRotAssAufDemTisch = True
    End If
    'liegt der RoterK�nig auf dem Tisch?                                            'Gerbing 16.10.2017 19.12.2017
    blnRoterK�nigAufDemTisch = False
    If (SpielerGespielt.Visible And SpielerGespielt.CardType = 2 And SpielerGespielt.CardValue = 7) Or _
        (LinksGespielt.Visible And LinksGespielt.CardType = 2 And LinksGespielt.CardValue = 7) Or _
        (MitteGespielt.Visible And MitteGespielt.CardType = 2 And MitteGespielt.CardValue = 7) Or _
        (RechtsGespielt.Visible And RechtsGespielt.CardType = 2 And RechtsGespielt.CardValue = 7) Then
        blnRoterK�nigAufDemTisch = True
    End If
    If PlayerTurn = Links Then
        For I = 1 To UBound(LinksCards)
            If LinksCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            'Computer hat nicht die geforderte Farbe, soll vorrangig RoterK�nig oder RotAss abwerfen    'Gerbing 09.12.2017
            'Suche RoterK�nig oder RotAss, sonst die h�chste                                            'Gerbing 09.12.2017
            H�chsterWert = 0
            For I = 1 To UBound(LinksCards)
                If LinksCards(I).cValue = 7 And LinksCards(I).cType = 2 Then                'RoterK�nig
                    H�chstePosition = I
                    SucheBesteKarteRoterK�nigMitspieler = H�chstePosition
                    Exit Function
                Else
                    If LinksCards(I).cValue = 8 And LinksCards(I).cType = 2 Then            'RotAss Gerbing 09.12.2017
                        H�chstePosition = I
                        SucheBesteKarteRoterK�nigMitspieler = H�chstePosition
                        Exit Function
                    Else
                        If LinksCards(I).cValue >= H�chsterWert Then
                            H�chsterWert = LinksCards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                End If
            Next I
            SucheBesteKarteRoterK�nigMitspieler = H�chstePosition
            Exit Function
        Else
            'es muss zugegeben werden
            '(Wenn der RoterK�nig noch nicht gespielt wurde und nicht auf dem Tisch liegt
            'und ich bin der 4.Zugeber(3 Karten liegen auf dem Tisch))
            'oder RotAss liegt auf dem Tisch -> VorbehandlungRoterK�nig(die h�chste Rot zugeben)'Gerbing 08.12.2017
            If blnRoterK�nigAufDemTisch = False And blnMitteWarDran = True And blnRechtsWarDran = True And blnSpielerWarDran = True Then
                SucheBesteKarteRoterK�nigMitspieler = VorbehandlungRoterK�nig(LinksCards, PlayerTurn, AKType)  'Gerbing 16.10.2017
                If SucheBesteKarteRoterK�nigMitspieler <> 0 Then                        'Gerbing 28.11.2017
                    Exit Function
                End If
            End If
            'sonst, Farbe ist vorhanden, es muss zugegeben werden, ich bin nicht der 4.Zugeber
            'wenn zugegeben werden muss und es ist eine h�here Karte gespielt als RoterK�nig, dann den RoterK�nig zugeben, falls ich ihn habe
            NiedrigsterWert = 8                                                         '8=Ass=Startwert
            K�nigPosition = 0
            If blnRotAssAufDemTisch = True And AKType = 2 Then                                          'Gerbing 08.12.2017 19.12.2017
                'Rot muss zugegeben werden es ist eine h�here Karte auf dem Tisch als RoterK�nig
                For I = 1 To UBound(LinksCards)
                    If LinksCards(I).cValue = 7 And LinksCards(I).cType = 2 Then         'RoterK�nig
                        K�nigPosition = I
                    Else
                        If LinksCards(I).cType = 2 And LinksCards(I).cValue <= NiedrigsterWert Then 'Gerbing 08.12.2017
                            NiedrigsterWert = LinksCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                Next I
                If K�nigPosition <> 0 Then
                    'ich gebe den RoterK�nig zu
                    NiedrigstePosition = K�nigPosition
                End If
                SucheBesteKarteRoterK�nigMitspieler = NiedrigstePosition
                Exit Function
            End If
            '-----------------------------------------------------------------------------
            'wenn zugegeben werden muss und es ist eine niedrigere Karte gespielt als RoterK�nig, den RoterK�nig nur spielen
            'wenn er der letzte dieser Farbe ist
            NiedrigsterWert = 8                                                         '8=Ass=Startwert
            NiedrigstePosition = 0
            For I = 1 To UBound(LinksCards)
                If AKType = LinksCards(I).cType Then
                    If LinksCards(I).cValue = 7 And LinksCards(I).cType = 2 Then        'RoterK�nig
                        K�nigPosition = I
                    Else
                        If LinksCards(I).cValue <= NiedrigsterWert Then
                            NiedrigsterWert = LinksCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                End If
            Next I
            If NiedrigstePosition = 0 Then
                'es gibt nur noch den RoterK�nig
                NiedrigstePosition = K�nigPosition
            End If
            SucheBesteKarteRoterK�nigMitspieler = NiedrigstePosition
            Exit Function
        End If
    End If
'--------------------------------------------------------------------------------------------------
    If PlayerTurn = Mitte Then
        For I = 1 To UBound(MitteCards)
            If MitteCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            'Computer hat nicht die geforderte Farbe, soll vorrangig RoterK�nig oder RotAss abwerfen    'Gerbing 09.12.2017
            'Suche RoterK�nig oder RotAss, sonst die h�chste                                            'Gerbing 09.12.2017
            H�chsterWert = 0
            For I = 1 To UBound(MitteCards)
                If MitteCards(I).cValue = 7 And MitteCards(I).cType = 2 Then                'RoterK�nig
                    H�chstePosition = I
                    SucheBesteKarteRoterK�nigMitspieler = H�chstePosition
                    Exit Function
                Else
                    If MitteCards(I).cValue = 8 And MitteCards(I).cType = 2 Then            'RotAss Gerbing 09.12.2017
                        H�chstePosition = I
                        SucheBesteKarteRoterK�nigMitspieler = H�chstePosition
                        Exit Function
                    Else
                        If MitteCards(I).cValue >= H�chsterWert Then
                            H�chsterWert = MitteCards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                End If
            Next I
            SucheBesteKarteRoterK�nigMitspieler = H�chstePosition
            Exit Function
        Else
            '(Wenn der RoterK�nig noch nicht gespielt wurde und nicht auf dem Tisch liegt
            'und ich bin der 4.Zugeber(3 Karten liegen auf dem Tisch))
            'oder RotAss liegt auf dem Tisch -> VorbehandlungRoterK�nig(die h�chste Rot zugeben)'Gerbing 08.12.2017
            If blnRoterK�nigAufDemTisch = False And blnMitteWarDran = True And blnRechtsWarDran = True And blnSpielerWarDran = True Then
                SucheBesteKarteRoterK�nigMitspieler = VorbehandlungRoterK�nig(MitteCards, PlayerTurn, AKType)  'Gerbing 16.10.2017
                If SucheBesteKarteRoterK�nigMitspieler <> 0 Then                        'Gerbing 28.11.2017
                    Exit Function
                End If
            End If
            'sonst, Farbe ist vorhanden, es muss zugegeben werden, ich bin nicht der 4.Zugeber
            'wenn zugegeben werden muss und es ist eine h�here Karte gespielt als RoterK�nig, dann den RoterK�nig zugeben, falls ich ihn habe
            NiedrigsterWert = 8                                                         '8=Ass=Startwert
            K�nigPosition = 0
            If blnRotAssAufDemTisch = True And AKType = 2 Then                                          'Gerbing 08.12.2017 19.12.2017
                'es ist eine h�here Karte auf dem Tisch als RoterK�nig
                For I = 1 To UBound(MitteCards)
                    If MitteCards(I).cValue = 7 And MitteCards(I).cType = 2 Then         'RoterK�nig
                        K�nigPosition = I
                    Else
                        If MitteCards(I).cType = 2 And MitteCards(I).cValue <= NiedrigsterWert Then 'Gerbing 08.12.2017
                            NiedrigsterWert = MitteCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                Next I
                If K�nigPosition <> 0 Then
                    'ich gebe den RoterK�nig zu
                    NiedrigstePosition = K�nigPosition
                End If
                SucheBesteKarteRoterK�nigMitspieler = NiedrigstePosition
                Exit Function
            End If
            '-----------------------------------------------------------------------------
            'wenn zugegeben werden muss und es ist eine niedrigere Karte gespielt als RoterK�nig, den RoterK�nig nur spielen
            'wenn er der letzte dieser Farbe ist
            NiedrigsterWert = 8                                                         '8=Ass=Startwert
            NiedrigstePosition = 0
            For I = 1 To UBound(MitteCards)
                If AKType = MitteCards(I).cType Then
                    If MitteCards(I).cValue = 7 And MitteCards(I).cType = 2 Then        'RoterK�nig
                        K�nigPosition = I
                    Else
                        If MitteCards(I).cValue <= NiedrigsterWert Then
                            NiedrigsterWert = MitteCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                End If
            Next I
            If NiedrigstePosition = 0 Then
                'es gibt nur noch den RoterK�nig
                NiedrigstePosition = K�nigPosition
            End If
            SucheBesteKarteRoterK�nigMitspieler = NiedrigstePosition
            Exit Function
        End If
    End If
'--------------------------------------------------------------------------------------------------
    If PlayerTurn = Rechts Then
        For I = 1 To UBound(RechtsCards)
            If RechtsCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            'Computer hat nicht die geforderte Farbe, soll vorrangig RoterK�nig oder RotAss abwerfen    'Gerbing 09.12.2017
            'Suche RoterK�nig oder RotAss, sonst die h�chste                                            'Gerbing 09.12.2017
            H�chsterWert = 0
            For I = 1 To UBound(RechtsCards)
                If RechtsCards(I).cValue = 7 And RechtsCards(I).cType = 2 Then                'RoterK�nig
                    H�chstePosition = I
                    SucheBesteKarteRoterK�nigMitspieler = H�chstePosition
                    Exit Function
                Else
                    If RechtsCards(I).cValue = 8 And RechtsCards(I).cType = 2 Then            'RotAss Gerbing 09.12.2017
                        H�chstePosition = I
                        SucheBesteKarteRoterK�nigMitspieler = H�chstePosition
                        Exit Function
                    Else
                        If RechtsCards(I).cValue >= H�chsterWert Then
                            H�chsterWert = RechtsCards(I).cValue
                            H�chstePosition = I
                        End If
                    End If
                End If
            Next I
            SucheBesteKarteRoterK�nigMitspieler = H�chstePosition
            Exit Function
        Else
            '(Wenn der RoterK�nig noch nicht gespielt wurde und nicht auf dem Tisch liegt
            'und ich bin der 4.Zugeber(3 Karten liegen auf dem Tisch))
            'oder RotAss liegt auf dem Tisch -> VorbehandlungRoterK�nig(die h�chste Rot zugeben)'Gerbing 08.12.2017
            If blnRoterK�nigAufDemTisch = False And blnMitteWarDran = True And blnRechtsWarDran = True And blnSpielerWarDran = True Then
                SucheBesteKarteRoterK�nigMitspieler = VorbehandlungRoterK�nig(RechtsCards, PlayerTurn, AKType)  'Gerbing 16.10.2017
                If SucheBesteKarteRoterK�nigMitspieler <> 0 Then                        'Gerbing 28.11.2017
                    Exit Function
                End If
            End If
            'sonst, Farbe ist vorhanden, es muss zugegeben werden, ich bin nicht der 4.Zugeber
            'wenn zugegeben werden muss und es ist eine h�here Karte gespielt als RoterK�nig, dann den RoterK�nig zugeben, falls ich ihn habe
            NiedrigsterWert = 8                                                         '8=Ass=Startwert
            K�nigPosition = 0
            If blnRotAssAufDemTisch = True And AKType = 2 Then                                          'Gerbing 08.12.2017 19.12.2017
                'es ist eine h�here Karte auf dem Tisch als RoterK�nig
                For I = 1 To UBound(RechtsCards)
                    If RechtsCards(I).cValue = 7 And RechtsCards(I).cType = 2 Then         'RoterK�nig
                        K�nigPosition = I
                    Else
                        If RechtsCards(I).cType = 2 And RechtsCards(I).cValue <= NiedrigsterWert Then 'Gerbing 08.12.2017
                            NiedrigsterWert = RechtsCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                Next I
                If K�nigPosition <> 0 Then
                    'ich gebe den RoterK�nig zu
                    NiedrigstePosition = K�nigPosition
                End If
                SucheBesteKarteRoterK�nigMitspieler = NiedrigstePosition
                Exit Function
            End If
            '-----------------------------------------------------------------------------
            'wenn zugegeben werden muss und es ist eine niedrigere Karte gespielt als RoterK�nig, den RoterK�nig nur spielen
            'wenn er der letzte dieser Farbe ist
            NiedrigsterWert = 8                                                         '8=Ass=Startwert
            NiedrigstePosition = 0
            For I = 1 To UBound(RechtsCards)
                If AKType = RechtsCards(I).cType Then
                    If RechtsCards(I).cValue = 7 And RechtsCards(I).cType = 2 Then        'RoterK�nig
                        K�nigPosition = I
                    Else
                        If RechtsCards(I).cValue <= NiedrigsterWert Then
                            NiedrigsterWert = RechtsCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                End If
            Next I
            If NiedrigstePosition = 0 Then
                'es gibt nur noch den RoterK�nig
                NiedrigstePosition = K�nigPosition
            End If
            SucheBesteKarteRoterK�nigMitspieler = NiedrigstePosition
            Exit Function
        End If
    End If
End Function

Private Function SucheBesteKarteErsterUndLetzter()
'    Der Ausspieler soll beim ersten Ausspielen und beim letzten Ausspielen die niedrigste Karte ausspielen wie bei SucheBesteKarteRotOderSticheAusspieler_N2
'    Gerbing 17.10.2017 Intelligenzverbesserung bei den letzen drei Karten die niedrigste ausspielen
'    Der Ausspieler soll ab 2. Stich entweder seine einzige Farbe oder seine h�chste Karte ausspielen.
'    NachbehandlungAusspieler pr�ft, ob eine Karte gespielt werden soll die kein anderer mehr hat

'    Wenn der Mitspieler zugeben muss spielt er seine h�chste Karte
'    Wenn der Mitspieler abwerfen kann, wirft er seine h�chste ab oder seine einzige wenn es nicht gerade die 7 ist(NachbehandlungAbwerfen)
    Dim I As Integer
    Dim NiedrigsterWert As Integer
    Dim NiedrigstePosition As Integer
    Dim H�chsterWert As Integer
    Dim H�chstePosition As Integer
    Dim AKType As Integer                                                   'AusgespielteKarteType
    Dim AKValue As Integer                                                   'AusgespielteKarteValue
    Dim blnFarbeVorhanden As Boolean

    If Ausspieler = PlayerTurn Then
        'Der Computer muss Ausspielen
        If AKIndex <= 3 Then
            'hier ist es das erste Ausspielen
            SucheBesteKarteErsterUndLetzter = SucheBesteKarteRotOderSticheAusspieler_N2
            Exit Function
        End If
        'If AKIndex >= 29 And AKIndex <= 32 Then
        If AKIndex >= 20 Then                                               'Gerbing 17.10.2017
            'hier ist es das vorvorletzte, vorletzte oder letzte Ausspielen
            SucheBesteKarteErsterUndLetzter = SucheBesteKarteRotOderSticheAusspieler_N2
            Exit Function
        End If
        'hier ist es das zweite bis f�nfte Aussspielen
        H�chsterWert = 0
        If PlayerTurn = Links Then
            For I = 1 To UBound(LinksCards)
                If LinksCards(I).cValue >= H�chsterWert Then
                    H�chsterWert = LinksCards(I).cValue
                    H�chstePosition = I
                End If
            Next I
            SucheBesteKarteErsterUndLetzter = NachbehandlungAusspieler(LinksCards, H�chstePosition)
            Exit Function
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To UBound(MitteCards)
                If MitteCards(I).cValue >= H�chsterWert Then
                    H�chsterWert = MitteCards(I).cValue
                    H�chstePosition = I
                End If
            Next I
            SucheBesteKarteErsterUndLetzter = NachbehandlungAusspieler(MitteCards, H�chstePosition)
            Exit Function
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To UBound(RechtsCards)
                If RechtsCards(I).cValue >= H�chsterWert Then
                    H�chsterWert = RechtsCards(I).cValue
                    H�chstePosition = I
                End If
            Next I
            SucheBesteKarteErsterUndLetzter = NachbehandlungAusspieler(RechtsCards, H�chstePosition)
            Exit Function
        End If
    End If
    '---------------------------------------------------------------------------------------------------------------------
    'Der Computer muss zugeben oder abwerfen
    blnFarbeVorhanden = False
    If Ausspieler = Links Then
        AKType = LinksGespielt.CardType
        AKValue = LinksGespielt.CardValue
    ElseIf Ausspieler = Mitte Then
        AKType = MitteGespielt.CardType
        AKValue = MitteGespielt.CardValue
    ElseIf Ausspieler = Rechts Then
        AKType = RechtsGespielt.CardType
        AKValue = RechtsGespielt.CardValue
    ElseIf Ausspieler = Spieler Then
        AKType = SpielerGespielt.CardType
        AKValue = SpielerGespielt.CardValue
    End If
'    Wenn der Mitspieler zugeben muss spielt er beim ersten Stich seine niedrigste Karte, ab 2.Stich seine h�chste Karte
'    Wenn der Mitspieler abwerfen kann, wirft er seine h�chste ab oder seine einzige wenn es nicht gerade die 7 ist(NachbehandlungAbwerfen)

    If PlayerTurn = Links Then
        For I = 1 To UBound(LinksCards)
            If LinksCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            'h�chste Karte abwerfen
            H�chsterWert = 0
            For I = 1 To UBound(LinksCards)
                'If AKType = LinksCards(i).cType Then
                    If LinksCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = LinksCards(I).cValue
                        H�chstePosition = I
                    End If
                'End If
            Next I
            SucheBesteKarteErsterUndLetzter = NachbehandlungAbwerfen(LinksCards, H�chstePosition)
            Exit Function
        Else
            'Farbe ist vorhanden
            'Wenn der Mitspieler zugeben muss spielt er beim ersten Stich seine niedrigste Karte, ab 2.Stich seine h�chste Karte
            If AKIndex <= 3 Then
                'hier ist es das erste Ausspielen
                NiedrigsterWert = 8                                             '8=Ass=Startwert
                For I = 1 To UBound(LinksCards)
                    If AKType = LinksCards(I).cType Then
                        If LinksCards(I).cValue <= NiedrigsterWert Then
                            NiedrigsterWert = LinksCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                Next I
                SucheBesteKarteErsterUndLetzter = NiedrigstePosition
                Exit Function
            End If
            'hier ab 2.Stich
            H�chsterWert = 0
            For I = 1 To UBound(LinksCards)
                If AKType = LinksCards(I).cType Then
                    If LinksCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = LinksCards(I).cValue
                        H�chstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteErsterUndLetzter = H�chstePosition
            Exit Function
        End If
    End If
'--------------------------------------------------------------------------------------------------
    If PlayerTurn = Mitte Then
        For I = 1 To UBound(MitteCards)
            If MitteCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            'h�chste Karte abwerfen
            H�chsterWert = 0
            For I = 1 To UBound(MitteCards)
                'If AKType = MitteCards(i).cType Then
                    If MitteCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = MitteCards(I).cValue
                        H�chstePosition = I
                    End If
                'End If
            Next I
            SucheBesteKarteErsterUndLetzter = NachbehandlungAbwerfen(MitteCards, H�chstePosition)
            Exit Function
        Else
            'Farbe ist vorhanden
            'Wenn der Mitspieler zugeben muss spielt er beim ersten Stich seine niedrigste Karte, ab 2.Stich seine h�chste Karte
            If AKIndex <= 3 Then
                'hier ist es das erste Ausspielen
                NiedrigsterWert = 8                                             '8=Ass=Startwert
                For I = 1 To UBound(MitteCards)
                    If AKType = MitteCards(I).cType Then
                        If MitteCards(I).cValue <= NiedrigsterWert Then
                            NiedrigsterWert = MitteCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                Next I
                SucheBesteKarteErsterUndLetzter = NiedrigstePosition
                Exit Function
            End If
            'hier ab 2.Stich
            H�chsterWert = 0
            For I = 1 To UBound(MitteCards)
                If AKType = MitteCards(I).cType Then
                    If MitteCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = MitteCards(I).cValue
                        H�chstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteErsterUndLetzter = H�chstePosition
            Exit Function
        End If
    End If
'--------------------------------------------------------------------------------------------------
    If PlayerTurn = Rechts Then
        For I = 1 To UBound(RechtsCards)
            If RechtsCards(I).cType = AKType Then
                blnFarbeVorhanden = True
            End If
        Next I
        If blnFarbeVorhanden = False Then
            'Farbe nicht vorhanden, es darf abgeworfen werden
            'h�chste Karte abwerfen
            H�chsterWert = 0
            For I = 1 To UBound(RechtsCards)
                'If AKType = RechtsCards(i).cType Then
                    If RechtsCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = RechtsCards(I).cValue
                        H�chstePosition = I
                    End If
                'End If
            Next I
            SucheBesteKarteErsterUndLetzter = NachbehandlungAbwerfen(RechtsCards, H�chstePosition)
            Exit Function
        Else
            'Farbe ist vorhanden
            'Wenn der Mitspieler zugeben muss spielt er beim ersten Stich seine niedrigste Karte, ab 2.Stich seine h�chste Karte
            If AKIndex <= 3 Then
                'hier ist es das erste Ausspielen
                NiedrigsterWert = 8                                             '8=Ass=Startwert
                For I = 1 To UBound(RechtsCards)
                    If AKType = RechtsCards(I).cType Then
                        If RechtsCards(I).cValue <= NiedrigsterWert Then
                            NiedrigsterWert = RechtsCards(I).cValue
                            NiedrigstePosition = I
                        End If
                    End If
                Next I
                SucheBesteKarteErsterUndLetzter = NiedrigstePosition
                Exit Function
            End If
            'hier ab 2.Stich
            H�chsterWert = 0
            For I = 1 To UBound(RechtsCards)
                If AKType = RechtsCards(I).cType Then
                    If RechtsCards(I).cValue >= H�chsterWert Then
                        H�chsterWert = RechtsCards(I).cValue
                        H�chstePosition = I
                    End If
                End If
            Next I
            SucheBesteKarteErsterUndLetzter = H�chstePosition
            Exit Function
        End If
    End If

End Function

Private Function SucheBesteKarteLorumAusspieler()
    'Der Ausspieler darf Asse ablegen
    'An auf dem Tisch liegende Karten d�rfen abw�rts weitere Karten angelegt werden
    'Es soll versucht werden eine aufsteigende Kartenfolge einer Farbe mit maximal 4 Karten zu erzeugen
    '   dabei ist es m�glich dass andere Mitspieler Karten spielen m�ssen die in der Kartenfolge fehlen
    '   Der Computer kann alle Karten durchsuchen, dann weis er, wer diese Karte hat
    'Wenn der Ausspieler nicht mehr weiter kann, kommt als n�chster Spieler der dran, der die letzte Karte einer vom
    'Ausspieler begonnenen Kartenfolge hinlegt
    'Die letzte Karte einer begonnenen Kartenfolge muss nicht die 4. sein, wenn diese bereits auf dem Tisch liegt
    'Lorum ist zu Ende, wenn einer der Spieler seine letzte Karte abgelegt hat

    Dim BestePosition As Integer
    Dim rc As Integer
    Dim Index As Integer

    'Der Ausspieler darf Asse ablegen
    'Hier suche ich Asse
    blnEingesammelt = True
    rc = LorumAssAblegen()
    If rc = 1 Then
'        gstrMsg = "Wer ein Ass ablegt darf weiter ausspielen"                              'Gerbing 15.08.2017
'        frmMsgBox.Show 1
        'Mittels Timer veranlassen, dass nach dem Ass durch das Programm weitergespielt wird
        TimerWeiterspielenNachAss.Enabled = True                                            'Gerbing 20.10.2017
        Exit Function                            'rc=1=Ass gefunden
    End If
    '-----------------------------------------------------------------------------------
    'Der Ausspieler darf an auf dem Tisch liegende Karten abw�rts weitere Karten ablegen
    'Hier suche ich Ablegekarten
    'MsgBox "Der letzte Ausspieler darf weiter ausspielen"
    Do                                                      'Schleife es d�rfen mehrere Karten abgelegt werden
        rc = LorumAbw�rtsAblegen()
        If rc = 0 Then Exit Do                              'keine Ablegekarte gefunden verlasse die Schleife
    Loop
    Do
        '-----------------------------------------------------------------------------------
        'hier versuche ich eine aufsteigende Kartenfolge einer Farbe mit maximal 4 Karten zu erzeugen
        'dabei ist es m�glich dass andere Mitspieler Karten spielen m�ssen die in der Kartenfolge fehlen
        'Der Computer weis, wer diese Karte hat
        rc = LorumKartenFolge(True)    'True=Loop benutzen
        If rc = 0 Then
    '        gstrMsg = "Ende der 4er-Folge erreicht" & vbNewLine        'Gerbing 15.08.2017
    '        gstrMsg = gstrMsg & "oder keine 4er-Folge erzeugbar"
    '        frmMsgBox.Show 1
            Exit Do                              'keine Ablegekarte erzeugbar verlasse die Schleife
        End If
        If blnLorumSpielende = True Then Exit Function
    Loop
    If blnLorumSpielende = True Then Exit Function
    'wenn sonst nichts mehr geht, eine Karte ablegen
    'Wenn der Ausspieler nicht mehr weiter kann, kommt als n�chster Spieler der dran, der die letzte Karte einer vom
    'Ausspieler begonnenen Kartenfolge hinlegt. Der Computer weis, wer die ben�tigte Karte hat
    'Zum Weiterspielen muss auf die Karten des vom Computer bestimmten Spielers geklickt werden
        Index = 0                                               'erste Karte nehmen
        'index = 2                                               'erste Karte nehmen
        blnErsteKommtVonSpieler = False
        If PlayerTurn = Links Then
            'MsgBox "Links legt eine einzelne Karte ab"
            blnErsteKommtVonLinks = True
            Call LorumEineKarteHinlegen(Index)
        ElseIf PlayerTurn = Mitte Then
            'MsgBox "Mitte legt eine einzelne Karte ab"
            blnErsteKommtVonMitte = True
            Call LorumEineKarteHinlegen(Index)
        ElseIf PlayerTurn = Rechts Then
            'MsgBox "Rechts legt eine einzelne Karte ab"
            blnErsteKommtVonRechts = True
            Call LorumEineKarteHinlegen(Index)
        End If
End Function

Private Function LorumKartenFolge(blnLoopBenutzen As Boolean)
    'rc = 0 keine Kartenfolge erzeugbar
    'rc = 1 Kartenfolge erzeugbar
    'Der Computer durchsucht seine Karten Spieler Links Mitte Rechts
    'und versucht eine aufsteigende Kartenfolge einer Farbe mit maximal 4 Karten zu erzeugen
    'dabei ist es m�glich dass andere Mitspieler Karten spielen m�ssen die in der Kartenfolge fehlen
    'Der Computer kann alle Karten durchsuchen, dann weis er, wer diese Karte hat
    
    Dim strFarbe As String
    Dim strWert As String
    Dim AP As Integer                                                   'AP=Auflegeposition
    Dim I As Integer
    Dim n As Integer
    Dim blnKartenFolgeErzeugbar As Boolean
    Dim blnN�chsteGefunden As Boolean
    
    If blnLoopBenutzen = True Then
        'Loop=True
        'hier die Kartenfolge ab Startkarte auflegen
        LorumAufgelegtZ�hler = 1
        PlayerTurnRoserl = PlayerTurn
        blnKartenFolgeErzeugbar = False
        
        If PlayerTurn = Spieler Then
            For I = 1 To UBound(SpielerCards)
                LorumStartWert = SpielerCards(I).cValue
                LorumStartPosition = I
                If LorumStartWert < 6 Then                                           '6=Ober
                    'durchsuche alle anderen Karten au�er der mit LorumStartWert, ob sie LorumStartWert+3 enth�lt und die gleiche Farbe hat
                    'und cGespieltVon <> 9
                    'Ober oder h�her zu untersuchen ist witzlos
                    For n = 1 To UBound(SpielerCards)
                        If SpielerCards(n).cValue <> LorumStartWert Then
                            If SpielerCards(n).cValue = LorumStartWert + 3 And SpielerCards(I).cType = SpielerCards(n).cType And SpielerCards(n).cGespieltVon <> 9 Then
                                LorumAuflegeFarbe = SpielerCards(n).cType
                                blnKartenFolgeErzeugbar = True
                                Exit For
                            End If
                        End If
                    Next n
                    If blnKartenFolgeErzeugbar = True Then Exit For
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To UBound(LinksCards)
                LorumStartWert = LinksCards(I).cValue
                LorumStartPosition = I
                If LorumStartWert < 6 Then                                           '6=Ober
                    'durchsuche alle anderen Karten au�er der mit LorumStartWert, ob sie LorumStartWert+3 enth�lt und die gleiche Farbe hat
                    'und cGespieltVon <> 9
                    'Ober oder h�her zu untersuchen ist witzlos
                    For n = 1 To UBound(LinksCards)
                        If LinksCards(n).cValue <> LorumStartWert Then
                            If LinksCards(n).cValue = LorumStartWert + 3 And LinksCards(I).cType = LinksCards(n).cType And LinksCards(n).cGespieltVon <> 9 Then
                                LorumAuflegeFarbe = LinksCards(n).cType
                                blnKartenFolgeErzeugbar = True
                                Exit For
                            End If
                        End If
                    Next n
                    If blnKartenFolgeErzeugbar = True Then Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To UBound(MitteCards)
                LorumStartWert = MitteCards(I).cValue
                LorumStartPosition = I
                If LorumStartWert < 6 Then                                           '6=Ober
                    'durchsuche alle anderen Karten au�er der mit LorumStartWert, ob sie LorumStartWert+3 enth�lt und die gleiche Farbe hat
                    'und cGespieltVon <> 9
                    'Ober oder h�her zu untersuchen ist witzlos
                    For n = 1 To UBound(MitteCards)
                        If MitteCards(n).cValue <> LorumStartWert Then
                            If MitteCards(n).cValue = LorumStartWert + 3 And MitteCards(I).cType = MitteCards(n).cType And MitteCards(n).cGespieltVon <> 9 Then
                                LorumAuflegeFarbe = MitteCards(n).cType
                                blnKartenFolgeErzeugbar = True
                                Exit For
                            End If
                        End If
                    Next n
                    If blnKartenFolgeErzeugbar = True Then Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To UBound(RechtsCards)
                LorumStartWert = RechtsCards(I).cValue
                LorumStartPosition = I
                If LorumStartWert < 6 Then                                           '6=Ober
                    'durchsuche alle anderen Karten au�er der mit LorumStartWert, ob sie LorumStartWert+3 enth�lt und die gleiche Farbe hat
                    'und cGespieltVon <> 9
                    'Ober oder h�her zu untersuchen ist witzlos
                    For n = 1 To UBound(RechtsCards)
                        If RechtsCards(n).cValue <> LorumStartWert Then
                            If RechtsCards(n).cValue = LorumStartWert + 3 And RechtsCards(I).cType = RechtsCards(n).cType And RechtsCards(n).cGespieltVon <> 9 Then
                                LorumAuflegeFarbe = RechtsCards(n).cType
                                blnKartenFolgeErzeugbar = True
                                Exit For
                            End If
                        End If
                    Next n
                    If blnKartenFolgeErzeugbar = True Then Exit For
                End If
            Next I
        End If

        
        If blnKartenFolgeErzeugbar = False Then
            LorumKartenFolge = 0                                            'Kartenfolge nicht erzeugbar
            Exit Function
        End If
    End If
    If blnLoopBenutzen = False Then                                         'Gerbing 26.09.2017
        '----------------------------------------
        'blnLoopBenutzen=False
        'Kartenfolge ist erzeugbar, oder es soll anstelle einer 4-er-Folge eine einzelne Karte abgelegt werden
        'das ist der Fall bei LorumEineKarteHinlegen
        LorumKartenFolge = 1                                                'Kartenfolge erzeugbar
    
        Do
            Call WelchenAblegestapelBenutzen(LorumAuflegeFarbe)
            If LorumAuflegeFarbe = 1 Then                                        '1=Schell
                If cdRoserlSchellA.Count <= 4 Then
                    If blnStapelBBenutzen = False Then
                        AP = cdRoserlSchellA.Count - 1
                        ZielPunktLeftLorum = cdRoserlSchellA(AP).Left           'Gerbing 25.09.2017
                        ZielPunktTopLorum = cdRoserlSchellA(AP).Top             'Gerbing 25.09.2017
                        Call MoveKarteLorum(LorumStartPosition - 1)             'Gerbing 25.09.2017
                        cdRoserlSchellA(AP).CardType = 1
                        cdRoserlSchellA(AP).CardValue = LorumStartWert
                        cdRoserlSchellA(AP).FaceMode = NachOben
                        cdRoserlSchellA(AP).Left = cdRoserlSchellA(0).Left + (CARD_SPACE / 2 * AP)
                        cdRoserlSchellA(AP).Top = cdRoserlSchellA(0).Top
                        cdRoserlSchellA(AP).Visible = True
                        On Error Resume Next                                'Gerbing 25.09.2017
                        Load cdRoserlSchellA(AP + 1)
                        On Error GoTo 0
                    End If
                End If
                If blnStapelBBenutzen = True Then
                    AP = cdRoserlSchellB.Count - 1
                    ZielPunktLeftLorum = cdRoserlSchellB(AP).Left           'Gerbing 25.09.2017
                    ZielPunktTopLorum = cdRoserlSchellB(AP).Top             'Gerbing 25.09.2017
                    'Hier
                    Call MoveKarteLorum(LorumStartPosition - 1)             'Gerbing 25.09.2017
                    cdRoserlSchellB(AP).CardType = 1
                    cdRoserlSchellB(AP).CardValue = LorumStartWert
                    cdRoserlSchellB(AP).FaceMode = NachOben
                    cdRoserlSchellB(AP).Left = cdRoserlSchellB(0).Left + (CARD_SPACE / 2 * AP)
                    cdRoserlSchellB(AP).Top = cdRoserlSchellB(0).Top
                    cdRoserlSchellB(AP).Visible = True
                    On Error Resume Next                                'Gerbing 25.09.2017
                    Load cdRoserlSchellB(AP + 1)
                    On Error GoTo 0
                End If
            End If
            If LorumAuflegeFarbe = 2 Then                                        '2=Rot
                If cdRoserlRotA.Count <= 4 Then
                    If blnStapelBBenutzen = False Then
                        AP = cdRoserlRotA.Count - 1
                        ZielPunktLeftLorum = cdRoserlRotA(AP).Left          'Gerbing 25.09.2017
                        ZielPunktTopLorum = cdRoserlRotA(AP).Top            'Gerbing 25.09.2017
                        Call MoveKarteLorum(LorumStartPosition - 1)         'Gerbing 25.09.2017
                        cdRoserlRotA(AP).CardType = 2
                        cdRoserlRotA(AP).CardValue = LorumStartWert
                        cdRoserlRotA(AP).FaceMode = NachOben
                        cdRoserlRotA(AP).Left = cdRoserlRotA(0).Left + (CARD_SPACE / 2 * AP)
                        cdRoserlRotA(AP).Top = cdRoserlRotA(0).Top
                        cdRoserlRotA(AP).Visible = True
                        On Error Resume Next                                'Gerbing 25.09.2017
                        Load cdRoserlRotA(AP + 1)
                        On Error GoTo 0
                    End If
                End If
                If blnStapelBBenutzen = True Then
                    AP = cdRoserlRotB.Count - 1
                    ZielPunktLeftLorum = cdRoserlRotB(AP).Left          'Gerbing 25.09.2017
                    ZielPunktTopLorum = cdRoserlRotB(AP).Top            'Gerbing 25.09.2017
                    Call MoveKarteLorum(LorumStartPosition - 1)         'Gerbing 25.09.2017
                    cdRoserlRotB(AP).CardType = 2
                    cdRoserlRotB(AP).CardValue = LorumStartWert
                    cdRoserlRotB(AP).FaceMode = NachOben
                    cdRoserlRotB(AP).Left = cdRoserlRotB(0).Left + (CARD_SPACE / 2 * AP)
                    cdRoserlRotB(AP).Top = cdRoserlRotB(0).Top
                    cdRoserlRotB(AP).Visible = True
                    On Error Resume Next                                'Gerbing 25.09.2017
                    Load cdRoserlRotB(AP + 1)
                    On Error GoTo 0
                End If
            End If
            If LorumAuflegeFarbe = 3 Then                                        '3=Gr�n
                If cdRoserlGr�nA.Count <= 4 Then
                    If blnStapelBBenutzen = False Then
                        AP = cdRoserlGr�nA.Count - 1
                        ZielPunktLeftLorum = cdRoserlGr�nA(AP).Left             'Gerbing 25.09.2017
                        ZielPunktTopLorum = cdRoserlGr�nA(AP).Top               'Gerbing 25.09.2017
                        Call MoveKarteLorum(LorumStartPosition - 1)             'Gerbing 25.09.2017
                        cdRoserlGr�nA(AP).CardType = 3
                        cdRoserlGr�nA(AP).CardValue = LorumStartWert
                        cdRoserlGr�nA(AP).FaceMode = NachOben
                        cdRoserlGr�nA(AP).Left = cdRoserlGr�nA(0).Left + (CARD_SPACE / 2 * AP)
                        cdRoserlGr�nA(AP).Top = cdRoserlGr�nA(0).Top
                        cdRoserlGr�nA(AP).Visible = True
                        On Error Resume Next                                'Gerbing 25.09.2017
                        Load cdRoserlGr�nA(AP + 1)
                        On Error GoTo 0
                    End If
                End If
                If blnStapelBBenutzen = True Then
                    AP = cdRoserlGr�nB.Count - 1
                    ZielPunktLeftLorum = cdRoserlGr�nB(AP).Left             'Gerbing 25.09.2017
                    ZielPunktTopLorum = cdRoserlGr�nB(AP).Top               'Gerbing 25.09.2017
                    Call MoveKarteLorum(LorumStartPosition - 1)             'Gerbing 25.09.2017
                    cdRoserlGr�nB(AP).CardType = 3
                    cdRoserlGr�nB(AP).CardValue = LorumStartWert
                    cdRoserlGr�nB(AP).FaceMode = NachOben
                    cdRoserlGr�nB(AP).Left = cdRoserlGr�nB(0).Left + (CARD_SPACE / 2 * AP)
                    cdRoserlGr�nB(AP).Top = cdRoserlGr�nB(0).Top
                    cdRoserlGr�nB(AP).Visible = True
                    On Error Resume Next                                'Gerbing 25.09.2017
                    Load cdRoserlGr�nB(AP + 1)
                    On Error GoTo 0
                End If
            End If
            If LorumAuflegeFarbe = 4 Then                                        '4=Eichel
                If cdRoserlEichelA.Count <= 4 Then
                    If blnStapelBBenutzen = False Then
                        AP = cdRoserlEichelA.Count - 1
                        ZielPunktLeftLorum = cdRoserlEichelA(AP).Left               'Gerbing 25.09.2017
                        ZielPunktTopLorum = cdRoserlEichelA(AP).Top                 'Gerbing 25.09.2017
                        Call MoveKarteLorum(LorumStartPosition - 1)                 'Gerbing 25.09.2017
                        cdRoserlEichelA(AP).CardType = 4
                        cdRoserlEichelA(AP).CardValue = LorumStartWert
                        cdRoserlEichelA(AP).FaceMode = NachOben
                        cdRoserlEichelA(AP).Left = cdRoserlEichelA(0).Left + (CARD_SPACE / 2 * AP)
                        cdRoserlEichelA(AP).Top = cdRoserlEichelA(0).Top
                        cdRoserlEichelA(AP).Visible = True
                        On Error Resume Next                                'Gerbing 25.09.2017
                        Load cdRoserlEichelA(AP + 1)
                        On Error GoTo 0
                    End If
                End If
                If blnStapelBBenutzen = True Then
                    AP = cdRoserlEichelB.Count - 1
                    ZielPunktLeftLorum = cdRoserlEichelB(AP).Left               'Gerbing 25.09.2017
                    ZielPunktTopLorum = cdRoserlEichelB(AP).Top                 'Gerbing 25.09.2017
                    Call MoveKarteLorum(LorumStartPosition - 1)                 'Gerbing 25.09.2017
                    cdRoserlEichelB(AP).CardType = 4
                    cdRoserlEichelB(AP).CardValue = LorumStartWert
                    cdRoserlEichelB(AP).FaceMode = NachOben
                    cdRoserlEichelB(AP).Left = cdRoserlEichelB(0).Left + (CARD_SPACE / 2 * AP)
                    cdRoserlEichelB(AP).Top = cdRoserlEichelB(0).Top
                    cdRoserlEichelB(AP).Visible = True
                    On Error Resume Next                                'Gerbing 25.09.2017
                    Load cdRoserlEichelB(AP + 1)
                    On Error GoTo 0
                End If
            End If
            '--------------------------------------------------------------------------
            'Die abgelegte Karte aus Spieler oder Links oder Mitte oder Rechts entfernen
            'Call EntladenLorum(PlayerTurnRoserl, LorumStartPosition)
            Call EntladenLorum(PlayerTurnRoserl)                                'Gerbing 17.12.2017
            If blnLorumSpielende = True Then Exit Function
            
            If LorumAufgelegtZ�hler = 4 Then
                LorumAufgelegtZ�hler = 0
                LorumKartenFolge = 0                    'Kartenfolge nicht erzeugbar
                Exit Function
            End If
            '--------------------------------------------------------------------------
            'Wem geh�rt die n�chste Karte
            LorumStartWert = LorumStartWert + 1
            blnN�chsteGefunden = False
            If LorumStartWert = 9 Then                                          'Gerbing 05.10.2017
                GoTo N�chsteGefunden
            End If
            For n = 1 To UBound(LinksCards)
                If LinksCards(n).cValue = LorumStartWert And LinksCards(n).cType = LorumAuflegeFarbe And LinksCards(n).cGespieltVon <> 9 Then
                    LorumStartPosition = n                                      '30.06.2017 LorumStartPosition war falsch wenn UBound(MitteCards) falsch war
                    PlayerTurnRoserl = 1                                        '1=Links
                    blnN�chsteGefunden = True
                    If LinksCards(n).cType = 1 Then strFarbe = "Schell"
                    If LinksCards(n).cType = 2 Then strFarbe = "Rot"
                    If LinksCards(n).cType = 3 Then strFarbe = "Gr�n"
                    If LinksCards(n).cType = 4 Then strFarbe = "Eichel"
                    If LorumStartWert = 1 Then strWert = "Sieben"
                    If LorumStartWert = 2 Then strWert = "Acht"
                    If LorumStartWert = 3 Then strWert = "Neun"
                    If LorumStartWert = 4 Then strWert = "Zehn"
                    If LorumStartWert = 5 Then strWert = "Unter"
                    If LorumStartWert = 6 Then strWert = "Ober"
                    If LorumStartWert = 7 Then strWert = "K�nig"
                    If LorumStartWert = 8 Then strWert = "Ass"
'                    gstrMsg = "Links legt jetzt eine Karte ab" & vbNewLine
'                    gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
'                    gstrMsg = gstrMsg & "Wert=" & strWert
'                    frmMsgBox.Show 1
                    'F�r MoveKarteLorum den StartPunkt festlegen                'Gerbing 25.09.2017
                    strAbleger = "Links"                                        'Gerbing 25.09.2017
                    StartPunktXLorum = cdLinks(LorumStartPosition - 1).Left     'Gerbing 25.09.2017
                    StartPunktYLorum = cdLinks(LorumStartPosition - 1).Top      'Gerbing 25.09.2017
                    GoTo N�chsteGefunden
                End If
            Next n
            For n = 1 To UBound(MitteCards)
                If MitteCards(n).cValue = LorumStartWert And MitteCards(n).cType = LorumAuflegeFarbe And MitteCards(n).cGespieltVon <> 9 Then
                    LorumStartPosition = n                                      '30.06.2017 LorumStartPosition war falsch wenn UBound(MitteCards) falsch war
                    PlayerTurnRoserl = 2                                        '2=mitte
                    blnN�chsteGefunden = True
                    If MitteCards(n).cType = 1 Then strFarbe = "Schell"
                    If MitteCards(n).cType = 2 Then strFarbe = "Rot"
                    If MitteCards(n).cType = 3 Then strFarbe = "Gr�n"
                    If MitteCards(n).cType = 4 Then strFarbe = "Eichel"
                    If LorumStartWert = 1 Then strWert = "Sieben"
                    If LorumStartWert = 2 Then strWert = "Acht"
                    If LorumStartWert = 3 Then strWert = "Neun"
                    If LorumStartWert = 4 Then strWert = "Zehn"
                    If LorumStartWert = 5 Then strWert = "Unter"
                    If LorumStartWert = 6 Then strWert = "Ober"
                    If LorumStartWert = 7 Then strWert = "K�nig"
                    If LorumStartWert = 8 Then strWert = "Ass"
'                    gstrMsg = "Mitte legt jetzt eine Karte ab" & vbNewLine
'                    gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
'                    gstrMsg = gstrMsg & "Wert=" & strWert
'                    frmMsgBox.Show 1
                    'F�r MoveKarteLorum den StartPunkt festlegen                'Gerbing 25.09.2017
                    strAbleger = "Mitte"                                        'Gerbing 25.09.2017
                    StartPunktXLorum = cdMitte(LorumStartPosition - 1).Left     'Gerbing 25.09.2017
                    StartPunktYLorum = cdMitte(LorumStartPosition - 1).Top      'Gerbing 25.09.2017
                    GoTo N�chsteGefunden
                End If
            Next n
            For n = 1 To UBound(RechtsCards)
                If RechtsCards(n).cValue = LorumStartWert And RechtsCards(n).cType = LorumAuflegeFarbe And RechtsCards(n).cGespieltVon <> 9 Then
                    LorumStartPosition = n                                      '30.06.2017 LorumStartPosition war falsch wenn UBound(RechtsCards) falsch war
                    PlayerTurnRoserl = 3                                        '3=mitte
                    blnN�chsteGefunden = True
                    If RechtsCards(n).cType = 1 Then strFarbe = "Schell"
                    If RechtsCards(n).cType = 2 Then strFarbe = "Rot"
                    If RechtsCards(n).cType = 3 Then strFarbe = "Gr�n"
                    If RechtsCards(n).cType = 4 Then strFarbe = "Eichel"
                    If LorumStartWert = 1 Then strWert = "Sieben"
                    If LorumStartWert = 2 Then strWert = "Acht"
                    If LorumStartWert = 3 Then strWert = "Neun"
                    If LorumStartWert = 4 Then strWert = "Zehn"
                    If LorumStartWert = 5 Then strWert = "Unter"
                    If LorumStartWert = 6 Then strWert = "Ober"
                    If LorumStartWert = 7 Then strWert = "K�nig"
                    If LorumStartWert = 8 Then strWert = "Ass"
'                    gstrMsg = "Rechts legt jetzt eine Karte ab" & vbNewLine
'                    gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
'                    gstrMsg = gstrMsg & "Wert=" & strWert
'                    frmMsgBox.Show 1
                    'F�r MoveKarteLorum den StartPunkt festlegen                'Gerbing 25.09.2017
                    strAbleger = "Rechts"                                       'Gerbing 25.09.2017
                    StartPunktXLorum = cdRechts(LorumStartPosition - 1).Left    'Gerbing 25.09.2017
                    StartPunktYLorum = cdRechts(LorumStartPosition - 1).Top     'Gerbing 25.09.2017
                    GoTo N�chsteGefunden
                End If
            Next n
            '---------------------------------------------------
            For n = 1 To UBound(SpielerCards)
                On Error Resume Next                            'Gerbing 29.06.2017
                SpielerCards(n).cValue = cdSpieler(n - 1).CardValue
                SpielerCards(n).cType = cdSpieler(n - 1).CardType
                On Error GoTo 0
                If SpielerCards(n).cValue = LorumStartWert And SpielerCards(n).cType = LorumAuflegeFarbe And SpielerCards(n).cGespieltVon <> 9 Then
                    LorumStartPosition = n                                      '30.06.2017 LorumStartPosition war falsch wenn UBound(SpielerCards) falsch war
                    PlayerTurnRoserl = 0                                        '0=Spieler
                    blnN�chsteGefunden = True
                    If SpielerCards(n).cType = 1 Then strFarbe = "Schell"
                    If SpielerCards(n).cType = 2 Then strFarbe = "Rot"
                    If SpielerCards(n).cType = 3 Then strFarbe = "Gr�n"
                    If SpielerCards(n).cType = 4 Then strFarbe = "Eichel"
                    If LorumStartWert = 1 Then strWert = "Sieben"
                    If LorumStartWert = 2 Then strWert = "Acht"
                    If LorumStartWert = 3 Then strWert = "Neun"
                    If LorumStartWert = 4 Then strWert = "Zehn"
                    If LorumStartWert = 5 Then strWert = "Unter"
                    If LorumStartWert = 6 Then strWert = "Ober"
                    If LorumStartWert = 7 Then strWert = "K�nig"
                    If LorumStartWert = 8 Then strWert = "Ass"
'                    gstrMsg = "Spieler legt jetzt eine Karte ab" & vbNewLine   'Gerbing 15.08.2017
'                    gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
'                    gstrMsg = gstrMsg & "Wert=" & strWert
'                    frmMsgBox.Show 1
                    'F�r MoveKarteLorum den StartPunkt festlegen                'Gerbing 25.09.2017
                    strAbleger = "Spieler"                                      'Gerbing 25.09.2017
                    StartPunktXLorum = cdSpieler(LorumStartPosition - 1).Left   'Gerbing 25.09.2017
                    StartPunktYLorum = cdSpieler(LorumStartPosition - 1).Top    'Gerbing 25.09.2017
                    GoTo N�chsteGefunden
                    Exit For
                End If
            Next n
N�chsteGefunden:
            If blnN�chsteGefunden = False Then
                LorumKartenFolge = 0                                            '0=dies war die letzte Karte der Kartenfolge
            Else
                LorumKartenFolge = 1                                            '1=es gibt eine weitere Karte in der Kartenfolge
            End If
            If LorumAufgelegtZ�hler = 0 Then
                LorumAufgelegtZ�hler = 1
            End If
            LorumAufgelegtZ�hler = LorumAufgelegtZ�hler + 1
            If blnLoopBenutzen = False Then Exit Function
        Loop
    End If
End Function

Private Function LorumAbw�rtsAblegen()
    'rc = 0 keine Anlegekarte gefunden
    'rc = 1 Anlegekarte gefunden
    'Der Computer durchsucht die Spieler Links Mitte Rechts Karten, ob an eine auf dem Tisch liegende Karte abw�rts angelegt werden kann
    
    Dim I As Integer
    Dim n As Integer
    Dim AP As Integer                                                   'AP=Anlegeposition
    Dim GKP As Integer                                                  'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
    Dim AnlegeWert As Integer
    Dim blnSchellAssAnlegekarteGefunden As Boolean
    Dim blnRotAssAnlegekarteGefunden As Boolean
    Dim blnGr�nAssAnlegekarteGefunden As Boolean
    Dim blnEichelAssAnlegekarteGefunden As Boolean
    Dim blnSchellAAnlegekarteGefunden As Boolean
    Dim blnSchellBAnlegekarteGefunden As Boolean
    Dim blnRotAAnlegekarteGefunden As Boolean
    Dim blnRotBAnlegekarteGefunden As Boolean
    Dim blnGr�nAAnlegekarteGefunden As Boolean
    Dim blnGr�nBAnlegekarteGefunden As Boolean
    Dim blnEichelAAnlegekarteGefunden As Boolean
    Dim blnEichelBAnlegekarteGefunden As Boolean
    Dim strFarbe As String
    Dim strWert As String

    AP = 0
    blnEichelBAnlegekarteGefunden = False
    For I = 0 To cdRoserlEichelB.Count - 1                              'f�r EichelB
        If cdRoserlEichelB(I).Visible = True Then                       'bei Visible=true ist sie aufgelegt
            AP = I + 1                                                  'AP sagt nur ob dieser Stapel sichtbar ist
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert ist eins unterhalb der am weitesten links liegenden karte
        AnlegeWert = cdRoserlEichelB(0).CardValue - 1
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 4 And SpielerCards(I).cGespieltVon <> 9 Then    '4=Eichel
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnEichelBAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnEichelBAnlegekarteGefunden = True
                        GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 4 And LinksCards(I).cGespieltVon <> 9 Then    '4=Eichel
                    blnEichelBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 4 And MitteCards(I).cGespieltVon <> 9 Then    '4=Eichel
                    blnEichelBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 4 And RechtsCards(I).cGespieltVon <> 9 Then    '4=Eichel
                    blnEichelBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
    End If
    If blnEichelBAnlegekarteGefunden = True Then
        GoTo KarteAuflegen
    End If
    '-----------------------------------------------
    AP = 0
    blnEichelAAnlegekarteGefunden = False
    For I = 0 To cdRoserlEichelA.Count - 1                              'f�r EichelA
        If cdRoserlEichelA(I).Visible = True Then                       'bei Visible=true ist sie aufgelegt
            AP = I + 1                                                  'AP sagt nur ob dieser Stapel sichtbar ist
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert ist eins unterhalb der am weitesten links liegenden karte
        AnlegeWert = cdRoserlEichelA(0).CardValue - 1
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 4 And SpielerCards(I).cGespieltVon <> 9 Then    '4=Eichel
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch angeklickt wurde.
                    'Wenn nein bleibt blnEichelAAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnEichelAAnlegekarteGefunden = True
                        GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 4 And LinksCards(I).cGespieltVon <> 9 Then    '4=Eichel
                    blnEichelAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 4 And MitteCards(I).cGespieltVon <> 9 Then    '4=Eichel
                    blnEichelAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 4 And RechtsCards(I).cGespieltVon <> 9 Then    '4=Eichel
                    blnEichelAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
    End If
    If blnEichelAAnlegekarteGefunden = True Then
        GoTo KarteAuflegen
    End If
    '-----------------------------------------------
    AP = 0
    blnGr�nAAnlegekarteGefunden = False
    For I = 0 To cdRoserlGr�nA.Count - 1                              'f�r Gr�nA
        If cdRoserlGr�nA(I).Visible = True Then                       'bei Visible=true ist sie aufgelegt
            AP = I + 1                                                  'AP sagt nur ob dieser Stapel sichtbar ist
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert ist eins unterhalb der am weitesten links liegenden karte
        AnlegeWert = cdRoserlGr�nA(0).CardValue - 1
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 3 And SpielerCards(I).cGespieltVon <> 9 Then    '3=Gr�n
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnGr�nAAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnGr�nAAnlegekarteGefunden = True
                        GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 3 And LinksCards(I).cGespieltVon <> 9 Then    '3=Gr�n
                    blnGr�nAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 3 And MitteCards(I).cGespieltVon <> 9 Then    '3=Gr�n
                    blnGr�nAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 3 And RechtsCards(I).cGespieltVon <> 9 Then    '3=Gr�n
                    blnGr�nAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
    End If
    If blnGr�nAAnlegekarteGefunden = True Then
        GoTo KarteAuflegen
    End If
    '-----------------------------------------------
    AP = 0
    blnGr�nBAnlegekarteGefunden = False
    For I = 0 To cdRoserlGr�nB.Count - 1                              'f�r Gr�nB
        If cdRoserlGr�nB(I).Visible = True Then                       'bei Visible=true ist sie aufgelegt
            AP = I + 1                                                  'AP sagt nur ob dieser Stapel sichtbar ist
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert ist eins unterhalb der am weitesten links liegenden karte
        AnlegeWert = cdRoserlGr�nB(0).CardValue - 1
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 3 And SpielerCards(I).cGespieltVon <> 9 Then    '3=Gr�n
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnGr�nBAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnGr�nBAnlegekarteGefunden = True
                        GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 3 And LinksCards(I).cGespieltVon <> 9 Then    '3=Gr�n
                    blnGr�nBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 3 And MitteCards(I).cGespieltVon <> 9 Then    '3=Gr�n
                    blnGr�nBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 3 And RechtsCards(I).cGespieltVon <> 9 Then    '3=Gr�n
                    blnGr�nBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
    End If
    If blnGr�nBAnlegekarteGefunden = True Then
        GoTo KarteAuflegen
    End If
    '-----------------------------------------------
    AP = 0
    blnRotAAnlegekarteGefunden = False
    For I = 0 To cdRoserlRotA.Count - 1                              'f�r RotA
        If cdRoserlRotA(I).Visible = True Then                       'bei Visible=true ist sie aufgelegt
            AP = I + 1                                                  'AP sagt nur ob dieser Stapel sichtbar ist
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert ist eins unterhalb der am weitesten links liegenden karte
        AnlegeWert = cdRoserlRotA(0).CardValue - 1
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 2 And SpielerCards(I).cGespieltVon <> 9 Then    '2=Rot
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnRotAAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnRotAAnlegekarteGefunden = True
                        GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 2 And LinksCards(I).cGespieltVon <> 9 Then    '2=Rot
                    blnRotAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 2 And MitteCards(I).cGespieltVon <> 9 Then    '2=Rot
                    blnRotAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 2 And RechtsCards(I).cGespieltVon <> 9 Then    '2=Rot
                    blnRotAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
    End If
    If blnRotAAnlegekarteGefunden = True Then
        GoTo KarteAuflegen
    End If
    '-----------------------------------------------
    AP = 0
    blnRotBAnlegekarteGefunden = False
    For I = 0 To cdRoserlRotB.Count - 1                              'f�r RotB
        If cdRoserlRotB(I).Visible = True Then                       'bei Visible=true ist sie aufgelegt
            AP = I + 1                                                  'AP sagt nur ob dieser Stapel sichtbar ist
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert ist eins unterhalb der am weitesten links liegenden karte
        AnlegeWert = cdRoserlRotB(0).CardValue - 1
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 2 And SpielerCards(I).cGespieltVon <> 9 Then    '2=Rot
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnRotBAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnRotBAnlegekarteGefunden = True
                        GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 2 And LinksCards(I).cGespieltVon <> 9 Then    '2=Rot
                    blnRotBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 2 And MitteCards(I).cGespieltVon <> 9 Then    '2=Rot
                    blnRotBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 2 And RechtsCards(I).cGespieltVon <> 9 Then    '2=Rot
                    blnRotBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
    End If
    If blnRotBAnlegekarteGefunden = True Then
        GoTo KarteAuflegen
    End If
'----------------------------------------------
    AP = 0
    blnSchellAAnlegekarteGefunden = False
    For I = 0 To cdRoserlSchellA.Count - 1                              'f�r SchellA
        If cdRoserlSchellA(I).Visible = True Then                       'bei Visible=true ist sie aufgelegt
            AP = I + 1                                                  'AP sagt nur ob dieser Stapel sichtbar ist
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert ist eins unterhalb der am weitesten links liegenden karte
        AnlegeWert = cdRoserlSchellA(0).CardValue - 1
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 1 And SpielerCards(I).cGespieltVon <> 9 Then    '1=Schell
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnSchellAAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnSchellAAnlegekarteGefunden = True
                        GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 1 And LinksCards(I).cGespieltVon <> 9 Then    '1=Schell
                    blnSchellAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 1 And MitteCards(I).cGespieltVon <> 9 Then    '1=Schell
                    blnSchellAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 1 And RechtsCards(I).cGespieltVon <> 9 Then    '1=Schell
                    blnSchellAAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
    End If
    If blnSchellAAnlegekarteGefunden = True Then
        GoTo KarteAuflegen
    End If
'----------------------------------------------
    AP = 0
    blnSchellBAnlegekarteGefunden = False
    For I = 0 To cdRoserlSchellB.Count - 1                              'f�r SchellB
        If cdRoserlSchellB(I).Visible = True Then                       'bei Visible=true ist sie aufgelegt
            AP = I + 1                                                  'AP sagt nur ob dieser Stapel sichtbar ist
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert ist eins unterhalb der am weitesten links liegenden karte
        AnlegeWert = cdRoserlSchellB(0).CardValue - 1
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 1 And SpielerCards(I).cGespieltVon <> 9 Then    '1=Schell
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnSchellBAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnSchellBAnlegekarteGefunden = True
                        GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 1 And LinksCards(I).cGespieltVon <> 9 Then    '1=Schell
                    blnSchellBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 1 And MitteCards(I).cGespieltVon <> 9 Then    '1=Schell
                    blnSchellBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 1 And RechtsCards(I).cGespieltVon <> 9 Then    '1=Schell
                    blnSchellBAnlegekarteGefunden = True
                    GKP = I                                                 'GKP=GespielteKartePosition wird f�rs Entladen gebraucht
                    Exit For
                End If
            Next I
        End If
    End If
    If blnSchellBAnlegekarteGefunden = True Then
        GoTo KarteAuflegen
    End If
    '-----------------------------------------------
    AP = 0
    blnSchellAssAnlegekarteGefunden = False
    For I = 0 To cdRoserlSchellAss.Count - 1                            'f�r SchellAss
        If cdRoserlSchellAss(I).Visible = True Then                     'bei Visible=true ist sie aufgelegt
            'finde die hinterste bereits aufgelegte Karte
            AP = I + 1
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert geht vom Ass=8 abw�rts vermindert um die AP
        AnlegeWert = 8 - AP
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 1 And SpielerCards(I).cGespieltVon <> 9 Then    '1=Schell
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnSchellAssAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnSchellAssAnlegekarteGefunden = True
                        GKP = I
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 1 And LinksCards(I).cGespieltVon <> 9 Then    '1=Schell
                    blnSchellAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 1 And MitteCards(I).cGespieltVon <> 9 Then    '1=Schell
                    blnSchellAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 1 And RechtsCards(I).cGespieltVon <> 9 Then    '1=Schell
                    blnSchellAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
    End If
    If blnSchellAssAnlegekarteGefunden = True Then GoTo KarteAuflegen
'----------------------------------------------------------------------
    AP = 0
    blnRotAssAnlegekarteGefunden = False
    For I = 0 To cdRoserlRotAss.Count - 1                            'f�r rotAss
        If cdRoserlRotAss(I).Visible = True Then                     'bei Visible=true ist sie aufgelegt
            'finde die hinterste bereits aufgelegte Karte
            AP = I + 1
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert geht vom Ass=8 abw�rts vermindert um die AP
        AnlegeWert = 8 - AP
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 2 And SpielerCards(I).cGespieltVon <> 9 Then    '2=rot
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnRotAssAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnRotAssAnlegekarteGefunden = True
                        GKP = I
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 2 And LinksCards(I).cGespieltVon <> 9 Then    '2=rot
                    blnRotAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 2 And MitteCards(I).cGespieltVon <> 9 Then    '2=rot
                    blnRotAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 2 And RechtsCards(I).cGespieltVon <> 9 Then    '2=rot
                    blnRotAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
    End If
    If blnRotAssAnlegekarteGefunden = True Then GoTo KarteAuflegen
'----------------------------------------------------------------------
    AP = 0
    blnGr�nAssAnlegekarteGefunden = False
    For I = 0 To cdRoserlGr�nAss.Count - 1                            'f�r gr�nAss
        If cdRoserlGr�nAss(I).Visible = True Then                     'bei Visible=true ist sie aufgelegt
            'finde die hinterste bereits aufgelegte Karte
            AP = I + 1
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert geht vom Ass=8 abw�rts vermindert um die AP
        AnlegeWert = 8 - AP
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 3 And SpielerCards(I).cGespieltVon <> 9 Then    '3=gr�n
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnGr�nAssAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnGr�nAssAnlegekarteGefunden = True
                        GKP = I
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 3 And LinksCards(I).cGespieltVon <> 9 Then    '3=gr�n
                    blnGr�nAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 3 And MitteCards(I).cGespieltVon <> 9 Then    '3=gr�n
                    blnGr�nAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 3 And RechtsCards(I).cGespieltVon <> 9 Then    '3=gr�n
                    blnGr�nAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
    End If
    If blnGr�nAssAnlegekarteGefunden = True Then GoTo KarteAuflegen
'----------------------------------------------------------------------
    AP = 0
    blnEichelAssAnlegekarteGefunden = False
    For I = 0 To cdRoserlEichelAss.Count - 1                            'f�r eichelAss
        If cdRoserlEichelAss(I).Visible = True Then                     'bei Visible=true ist sie aufgelegt
            'finde die hinterste bereits aufgelegte Karte
            AP = I + 1
        End If
    Next I
    If AP <> 0 Then
        'Der anzulegende Wert geht vom Ass=8 abw�rts vermindert um die AP
        AnlegeWert = 8 - AP
        If PlayerTurn = Spieler Then
            For I = 1 To 8
                If SpielerCards(I).cValue = AnlegeWert And SpielerCards(I).cType = 4 And SpielerCards(I).cGespieltVon <> 9 Then    '4=eichel
                    'Jetzt muss ich fragen ob die gefundene Ablegekarte auch vom Spieler angeklickt wurde.
                    'Wenn nein bleibt blnEichelAssAnlegekarteGefunden = False
                    'und es geht weiter mit dem Finden einer anderen Ablegekarte.
                    'Wenn ja geht es zu GoTo KarteAuflegen
                    If SpielerCards(I).cValue = LorumStartWert And SpielerCards(I).cType = LorumAuflegeFarbe Then           'Gerbing 26.09.2017
                        blnEichelAssAnlegekarteGefunden = True
                        GKP = I
                        Exit For
                    End If
                End If
            Next I
        End If
        If PlayerTurn = Links Then
            For I = 1 To 8
                If LinksCards(I).cValue = AnlegeWert And LinksCards(I).cType = 4 And LinksCards(I).cGespieltVon <> 9 Then    '4=eichel
                    blnEichelAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Mitte Then
            For I = 1 To 8
                If MitteCards(I).cValue = AnlegeWert And MitteCards(I).cType = 4 And MitteCards(I).cGespieltVon <> 9 Then    '4=eichel
                    blnEichelAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
        If PlayerTurn = Rechts Then
            For I = 1 To 8
                If RechtsCards(I).cValue = AnlegeWert And RechtsCards(I).cType = 4 And RechtsCards(I).cGespieltVon <> 9 Then    '4=eichel
                    blnEichelAssAnlegekarteGefunden = True
                    GKP = I
                    Exit For
                End If
            Next I
        End If
    End If
    If blnEichelAssAnlegekarteGefunden = True Then GoTo KarteAuflegen
    '-----------------------------------------------
    If blnSchellAssAnlegekarteGefunden = False And blnRotAssAnlegekarteGefunden = False And blnGr�nAssAnlegekarteGefunden = False _
        And blnEichelAssAnlegekarteGefunden = False And blnEichelAAnlegekarteGefunden = False And blnEichelBAnlegekarteGefunden = False _
        And blnGr�nAAnlegekarteGefunden = False And blnGr�nBAnlegekarteGefunden = False _
        And blnRotAAnlegekarteGefunden = False And blnRotBAnlegekarteGefunden = False _
        And blnSchellAAnlegekarteGefunden = False And blnSchellBAnlegekarteGefunden = False Then
        LorumAbw�rtsAblegen = 0
        Exit Function                                                   'keine Anlegekarte gefunden
    End If
KarteAuflegen:
    If PlayerTurn = Links Then
        strAbleger = "Links"
        StartPunktXLorum = cdLinks(GKP - 1).Left                                  'Gerbing 25.09.2017
        StartPunktYLorum = cdLinks(GKP - 1).Top                                   'Gerbing 25.09.2017
    ElseIf PlayerTurn = Mitte Then
        strAbleger = "Mitte"
        StartPunktXLorum = cdMitte(GKP - 1).Left                                  'Gerbing 25.09.2017
        StartPunktYLorum = cdMitte(GKP - 1).Top                                   'Gerbing 25.09.2017
    ElseIf PlayerTurn = Rechts Then
        strAbleger = "Rechts"
        StartPunktXLorum = cdRechts(GKP - 1).Left                                  'Gerbing 25.09.2017
        StartPunktYLorum = cdRechts(GKP - 1).Top                                   'Gerbing 25.09.2017
    ElseIf PlayerTurn = Spieler Then
        strAbleger = "Spieler"
        StartPunktXLorum = cdSpieler(GKP - 1).Left                                  'Gerbing 25.09.2017
        StartPunktYLorum = cdSpieler(GKP - 1).Top                                   'Gerbing 25.09.2017
    End If
    If blnEichelBAnlegekarteGefunden = True Or blnEichelAAnlegekarteGefunden = True Or blnEichelAssAnlegekarteGefunden = True Then
        strFarbe = "Eichel"
        LorumAuflegeFarbe = 4                                                       'Gerbing 17.12.2017
    End If
    If blnGr�nBAnlegekarteGefunden = True Or blnGr�nAAnlegekarteGefunden = True Or blnGr�nAssAnlegekarteGefunden = True Then
        strFarbe = "Gr�n"
        LorumAuflegeFarbe = 3                                                       'Gerbing 17.12.2017
    End If
    If blnRotBAnlegekarteGefunden = True Or blnRotAAnlegekarteGefunden = True Or blnRotAssAnlegekarteGefunden = True Then
        strFarbe = "Rot"
        LorumAuflegeFarbe = 2                                                       'Gerbing 17.12.2017
    End If
    If blnSchellBAnlegekarteGefunden = True Or blnSchellAAnlegekarteGefunden = True Or blnSchellAssAnlegekarteGefunden = True Then
        strFarbe = "Schell"
        LorumAuflegeFarbe = 1                                                       'Gerbing 17.12.2017
    End If
    If AnlegeWert = 1 Then strWert = "Sieben"
    If AnlegeWert = 2 Then strWert = "Acht"
    If AnlegeWert = 3 Then strWert = "Neun"
    If AnlegeWert = 4 Then strWert = "Zehn"
    If AnlegeWert = 5 Then strWert = "Unter"
    If AnlegeWert = 6 Then strWert = "Ober"
    If AnlegeWert = 7 Then strWert = "K�nig"
    If AnlegeWert = 8 Then strWert = "Ass"
    LorumStartWert = AnlegeWert                                                     'Gerbing 17.12.2017
    
'    gstrMsg = strAbleger & " legt jetzt eine Karte ab" & vbNewLine
'    gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
'    gstrMsg = gstrMsg & "Wert=" & strWert
'    frmMsgBox.Show 1
    LorumAbw�rtsAblegen = 1                                             'Anlegekarte gefunden
    '----------------------------------------------------------------------------------------
    'hier die Anlegekarte auflegen AP
    If blnEichelBAnlegekarteGefunden = True Then
        ZielPunktLeftLorum = cdRoserlEichelB(0).Left                'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlEichelB(0).Top                  'Gerbing 25.09.2017
        Call MoveKarteLorum(GKP - 1)                                'Gerbing 25.09.2017
        cdRoserlEichelB(0).CardType = 4
        cdRoserlEichelB(0).CardValue = AnlegeWert
        cdRoserlEichelB(0).FaceMode = NachOben
        cdRoserlEichelB(0).Left = cdRoserlEichelB(0).Left - (CARD_SPACE / 2)
        cdRoserlEichelB(0).Top = cdRoserlEichelB(0).Top
        AP = cdRoserlEichelB(0).Left                                    'Merke ganz linke Position
        If cdRoserlEichelB(0).CardValue + cdRoserlEichelB.Count = cdRoserlEichelB(cdRoserlEichelB.Count - 1).CardValue Then
            Load cdRoserlEichelB(cdRoserlEichelB.Count)                 'manchmal ist Load f�r eine zus�tzliche Karte n�tig
        End If
        For I = 0 To cdRoserlEichelB.Count - 1
            cdRoserlEichelB(I).Visible = True
            If I <> 0 Then
                cdRoserlEichelB(I).CardValue = AnlegeWert + I
                cdRoserlEichelB(I).Left = AP + (CARD_SPACE / 2 * I)
            End If
        Next I
    End If
    If blnEichelAAnlegekarteGefunden = True Then
        ZielPunktLeftLorum = cdRoserlEichelA(0).Left                'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlEichelA(0).Top                  'Gerbing 25.09.2017
        Call MoveKarteLorum(GKP - 1)                                'Gerbing 25.09.2017
        cdRoserlEichelA(0).CardType = 4
        cdRoserlEichelA(0).CardValue = AnlegeWert
        cdRoserlEichelA(0).FaceMode = NachOben
        cdRoserlEichelA(0).Left = cdRoserlEichelA(0).Left - (CARD_SPACE / 2)
        cdRoserlEichelA(0).Top = cdRoserlEichelA(0).Top
        AP = cdRoserlEichelA(0).Left                                    'Merke ganz linke Position
        If cdRoserlEichelA(0).CardValue + cdRoserlEichelA.Count = cdRoserlEichelA(cdRoserlEichelA.Count - 1).CardValue Then
            Load cdRoserlEichelA(cdRoserlEichelA.Count)                 'manchmal ist Load f�r eine zus�tzliche Karte n�tig
        End If
        For I = 0 To cdRoserlEichelA.Count - 1
            cdRoserlEichelA(I).Visible = True
            If I <> 0 Then
                cdRoserlEichelA(I).CardValue = AnlegeWert + I
                cdRoserlEichelA(I).Left = AP + (CARD_SPACE / 2 * I)
            End If
        Next I
    End If
    If blnGr�nBAnlegekarteGefunden = True Then
        ZielPunktLeftLorum = cdRoserlGr�nB(0).Left              'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlGr�nB(0).Top                'Gerbing 25.09.2017
        Call MoveKarteLorum(GKP - 1)                            'Gerbing 25.09.2017
        cdRoserlGr�nB(0).CardType = 3
        cdRoserlGr�nB(0).CardValue = AnlegeWert
        cdRoserlGr�nB(0).FaceMode = NachOben
        cdRoserlGr�nB(0).Left = cdRoserlGr�nB(0).Left - (CARD_SPACE / 2)
        cdRoserlGr�nB(0).Top = cdRoserlGr�nB(0).Top
        AP = cdRoserlGr�nB(0).Left                                    'Merke ganz linke Position
        If cdRoserlGr�nB(0).CardValue + cdRoserlGr�nB.Count = cdRoserlGr�nB(cdRoserlGr�nB.Count - 1).CardValue Then
            Load cdRoserlGr�nB(cdRoserlGr�nB.Count)                 'manchmal ist Load f�r eine zus�tzliche Karte n�tig
        End If
        For I = 0 To cdRoserlGr�nB.Count - 1
            cdRoserlGr�nB(I).Visible = True
            If I <> 0 Then
                cdRoserlGr�nB(I).CardValue = AnlegeWert + I
                cdRoserlGr�nB(I).Left = AP + (CARD_SPACE / 2 * I)
            End If
        Next I
    End If
    If blnGr�nAAnlegekarteGefunden = True Then
        ZielPunktLeftLorum = cdRoserlGr�nA(0).Left              'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlGr�nA(0).Top                'Gerbing 25.09.2017
        Call MoveKarteLorum(GKP - 1)                            'Gerbing 25.09.2017
        cdRoserlGr�nA(0).CardType = 3
        cdRoserlGr�nA(0).CardValue = AnlegeWert
        cdRoserlGr�nA(0).FaceMode = NachOben
        cdRoserlGr�nA(0).Left = cdRoserlGr�nA(0).Left - (CARD_SPACE / 2)
        cdRoserlGr�nA(0).Top = cdRoserlGr�nA(0).Top
        AP = cdRoserlGr�nA(0).Left                                    'Merke ganz linke Position
        If cdRoserlGr�nA(0).CardValue + cdRoserlGr�nA.Count = cdRoserlGr�nA(cdRoserlGr�nA.Count - 1).CardValue Then
            Load cdRoserlGr�nA(cdRoserlGr�nA.Count)                 'manchmal ist Load f�r eine zus�tzliche Karte n�tig
        End If
        For I = 0 To cdRoserlGr�nA.Count - 1
            cdRoserlGr�nA(I).Visible = True
            If I <> 0 Then
                cdRoserlGr�nA(I).CardValue = AnlegeWert + I
                cdRoserlGr�nA(I).Left = AP + (CARD_SPACE / 2 * I)
            End If
        Next I
    End If
    If blnRotBAnlegekarteGefunden = True Then
        ZielPunktLeftLorum = cdRoserlRotB(0).Left               'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlRotB(0).Top                 'Gerbing 25.09.2017
        Call MoveKarteLorum(GKP - 1)                            'Gerbing 25.09.2017
        cdRoserlRotB(0).CardType = 2
        cdRoserlRotB(0).CardValue = AnlegeWert
        cdRoserlRotB(0).FaceMode = NachOben
        cdRoserlRotB(0).Left = cdRoserlRotB(0).Left - (CARD_SPACE / 2)
        cdRoserlRotB(0).Top = cdRoserlRotB(0).Top
        AP = cdRoserlRotB(0).Left                                    'Merke ganz linke Position
        If cdRoserlRotB(0).CardValue + cdRoserlRotB.Count = cdRoserlRotB(cdRoserlRotB.Count - 1).CardValue Then
            Load cdRoserlRotB(cdRoserlRotB.Count)                 'manchmal ist Load f�r eine zus�tzliche Karte n�tig
        End If
        For I = 0 To cdRoserlRotB.Count - 1
            cdRoserlRotB(I).Visible = True
            If I <> 0 Then
                cdRoserlRotB(I).CardValue = AnlegeWert + I
                cdRoserlRotB(I).Left = AP + (CARD_SPACE / 2 * I)
            End If
        Next I
    End If
    If blnRotAAnlegekarteGefunden = True Then
        ZielPunktLeftLorum = cdRoserlRotA(0).Left               'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlRotA(0).Top                 'Gerbing 25.09.2017
        Call MoveKarteLorum(GKP - 1)                            'Gerbing 25.09.2017
        cdRoserlRotA(0).CardType = 2
        cdRoserlRotA(0).CardValue = AnlegeWert
        cdRoserlRotA(0).FaceMode = NachOben
        cdRoserlRotA(0).Left = cdRoserlRotA(0).Left - (CARD_SPACE / 2)
        cdRoserlRotA(0).Top = cdRoserlRotA(0).Top
        AP = cdRoserlRotA(0).Left                                    'Merke ganz linke Position
        If cdRoserlRotA(0).CardValue + cdRoserlRotA.Count = cdRoserlRotA(cdRoserlRotA.Count - 1).CardValue Then
            Load cdRoserlRotA(cdRoserlRotA.Count)                 'manchmal ist Load f�r eine zus�tzliche Karte n�tig
        End If
        For I = 0 To cdRoserlRotA.Count - 1
            cdRoserlRotA(I).Visible = True
            If I <> 0 Then
                cdRoserlRotA(I).CardValue = AnlegeWert + I
                cdRoserlRotA(I).Left = AP + (CARD_SPACE / 2 * I)
            End If
        Next I
    End If
    If blnSchellBAnlegekarteGefunden = True Then
        ZielPunktLeftLorum = cdRoserlSchellB(0).Left                'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlSchellB(0).Top                  'Gerbing 25.09.2017
        Call MoveKarteLorum(GKP - 1)                                'Gerbing 25.09.2017
        cdRoserlSchellB(0).CardType = 1
        cdRoserlSchellB(0).CardValue = AnlegeWert
        cdRoserlSchellB(0).FaceMode = NachOben
        cdRoserlSchellB(0).Left = cdRoserlSchellB(0).Left - (CARD_SPACE / 2)
        cdRoserlSchellB(0).Top = cdRoserlSchellB(0).Top
        AP = cdRoserlSchellB(0).Left                                    'Merke ganz linke Position
        If cdRoserlSchellB(0).CardValue + cdRoserlSchellB.Count = cdRoserlSchellB(cdRoserlSchellB.Count - 1).CardValue Then
            Load cdRoserlSchellB(cdRoserlSchellB.Count)                 'manchmal ist Load f�r eine zus�tzliche Karte n�tig
        End If
        For I = 0 To cdRoserlSchellB.Count - 1
            cdRoserlSchellB(I).Visible = True
            If I <> 0 Then
                cdRoserlSchellB(I).CardValue = AnlegeWert + I
                cdRoserlSchellB(I).Left = AP + (CARD_SPACE / 2 * I)
            End If
        Next I
    End If
    If blnSchellAAnlegekarteGefunden = True Then
        ZielPunktLeftLorum = cdRoserlSchellA(0).Left                'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlSchellA(0).Top                  'Gerbing 25.09.2017
        Call MoveKarteLorum(GKP - 1)                                'Gerbing 25.09.2017
        cdRoserlSchellA(0).CardType = 1
        cdRoserlSchellA(0).CardValue = AnlegeWert
        cdRoserlSchellA(0).FaceMode = NachOben
        cdRoserlSchellA(0).Left = cdRoserlSchellA(0).Left - (CARD_SPACE / 2)
        cdRoserlSchellA(0).Top = cdRoserlSchellA(0).Top
        AP = cdRoserlSchellA(0).Left                                    'Merke ganz linke Position
        If cdRoserlSchellA(0).CardValue + cdRoserlSchellA.Count = cdRoserlSchellA(cdRoserlSchellA.Count - 1).CardValue Then
            Load cdRoserlSchellA(cdRoserlSchellA.Count)                 'manchmal ist Load f�r eine zus�tzliche Karte n�tig
        End If
        For I = 0 To cdRoserlSchellA.Count - 1
            cdRoserlSchellA(I).Visible = True
            If I <> 0 Then
                cdRoserlSchellA(I).CardValue = AnlegeWert + I
                cdRoserlSchellA(I).Left = AP + (CARD_SPACE / 2 * I)
            End If
        Next I
    End If
    '--------------------------------------------------------------------
    If blnSchellAssAnlegekarteGefunden = True Then
        If strFarbe = "Schell" Then                               '1=Schell
            Load cdRoserlSchellAss(AP)
            ZielPunktLeftLorum = cdRoserlSchellAss(AP).Left             'Gerbing 25.09.2017
            ZielPunktTopLorum = cdRoserlSchellAss(AP).Top               'Gerbing 25.09.2017
            Call MoveKarteLorum(GKP - 1)                                'Gerbing 25.09.2017
            cdRoserlSchellAss(AP).CardType = 1
            cdRoserlSchellAss(AP).CardValue = AnlegeWert
            cdRoserlSchellAss(AP).FaceMode = NachOben
            cdRoserlSchellAss(AP).Left = cdRoserlSchellAss(0).Left - (CARD_SPACE / 2 * AP)
            cdRoserlSchellAss(AP).Top = cdRoserlSchellAss(0).Top
            For I = 0 To AP
                cdRoserlSchellAss(I).ZOrder 0
                cdRoserlSchellAss(I).Left = cdRoserlSchellAss(0).Left - (CARD_SPACE / 2 * I)
                cdRoserlSchellAss(I).Visible = True
            Next I
        End If
    End If
    If blnRotAssAnlegekarteGefunden = True Then
        If strFarbe = "Rot" Then                               '2=Rot
            Load cdRoserlRotAss(AP)
            ZielPunktLeftLorum = cdRoserlRotAss(AP).Left                'Gerbing 25.09.2017
            ZielPunktTopLorum = cdRoserlRotAss(AP).Top                  'Gerbing 25.09.2017
            Call MoveKarteLorum(GKP - 1)                                'Gerbing 25.09.2017
            cdRoserlRotAss(AP).CardType = 2
            cdRoserlRotAss(AP).CardValue = AnlegeWert
            cdRoserlRotAss(AP).FaceMode = NachOben
            cdRoserlRotAss(AP).Left = cdRoserlRotAss(0).Left - (CARD_SPACE / 2 * AP)
            cdRoserlRotAss(AP).Top = cdRoserlRotAss(0).Top
            For I = 0 To AP
                cdRoserlRotAss(I).ZOrder 0
                cdRoserlRotAss(I).Left = cdRoserlRotAss(0).Left - (CARD_SPACE / 2 * I)
                cdRoserlRotAss(I).Visible = True
            Next I
        End If
    End If
    If blnGr�nAssAnlegekarteGefunden = True Then
        If strFarbe = "Gr�n" Then                               '3=Gr�n
            Load cdRoserlGr�nAss(AP)
            ZielPunktLeftLorum = cdRoserlGr�nAss(AP).Left               'Gerbing 25.09.2017
            ZielPunktTopLorum = cdRoserlGr�nAss(AP).Top                 'Gerbing 25.09.2017
            Call MoveKarteLorum(GKP - 1)                                'Gerbing 25.09.2017
            cdRoserlGr�nAss(AP).CardType = 3
            cdRoserlGr�nAss(AP).CardValue = AnlegeWert
            cdRoserlGr�nAss(AP).FaceMode = NachOben
            cdRoserlGr�nAss(AP).Left = cdRoserlGr�nAss(0).Left - (CARD_SPACE / 2 * AP)
            cdRoserlGr�nAss(AP).Top = cdRoserlGr�nAss(0).Top
            For I = 0 To AP
                cdRoserlGr�nAss(I).ZOrder 0
                cdRoserlGr�nAss(I).Left = cdRoserlGr�nAss(0).Left - (CARD_SPACE / 2 * I)
                cdRoserlGr�nAss(I).Visible = True
            Next I
        End If
    End If
    If blnEichelAssAnlegekarteGefunden = True Then
        If strFarbe = "Eichel" Then                             '4=Eichel
            Load cdRoserlEichelAss(AP)
            ZielPunktLeftLorum = cdRoserlEichelAss(AP).Left             'Gerbing 25.09.2017
            ZielPunktTopLorum = cdRoserlEichelAss(AP).Top               'Gerbing 25.09.2017
            Call MoveKarteLorum(GKP - 1)                                'Gerbing 25.09.2017
            cdRoserlEichelAss(AP).CardType = 4
            cdRoserlEichelAss(AP).CardValue = AnlegeWert
            cdRoserlEichelAss(AP).FaceMode = NachOben
            cdRoserlEichelAss(AP).Left = cdRoserlEichelAss(0).Left - (CARD_SPACE / 2 * AP)
            cdRoserlEichelAss(AP).Top = cdRoserlEichelAss(0).Top
            For I = 0 To AP
                cdRoserlEichelAss(I).ZOrder 0
                cdRoserlEichelAss(I).Left = cdRoserlEichelAss(0).Left - (CARD_SPACE / 2 * I)
                cdRoserlEichelAss(I).Visible = True
            Next I
        End If
    End If
    '---------------------------------------------
    'Die Ablegekarte aus ComputerCards entfernen
    'Call EntladenLorum(PlayerTurn, GKP)
    Call EntladenLorum(PlayerTurn)                                      'Gerbing 17.12.2017
    End Function

Private Function LorumAssAblegen()
    'rc = 0 kein Ass gefunden
    'rc = 1 Ass gefunden
    'Der Ausspieler darf Asse ablegen
    'Hier suche ich Asse

    Dim I As Integer
    Dim blnAssGefunden As Boolean
    Dim strFarbe As String
    Dim strWert As String

    strWert = "Ass"
    blnAssGefunden = False

    If PlayerTurn = Links Then
        For I = 1 To UBound(LinksCards)
            If LinksCards(I).cValue = 8 And LinksCards(I).cGespieltVon <> 9 Then                    '8=Ass
                blnAssGefunden = True
                AssPosition = I
                If LinksCards(AssPosition).cType = 1 Then
                    strFarbe = "Schell"
                    LorumAuflegeFarbe = 1                                                           'Gerbing 17.12.2017
                End If
                If LinksCards(AssPosition).cType = 2 Then
                    strFarbe = "Rot"
                    LorumAuflegeFarbe = 2                                                           'Gerbing 17.12.2017
                End If
                If LinksCards(AssPosition).cType = 3 Then
                    strFarbe = "Gr�n"
                    LorumAuflegeFarbe = 3                                                           'Gerbing 17.12.2017
                End If
                If LinksCards(AssPosition).cType = 4 Then
                    strFarbe = "Eichel"
                    LorumAuflegeFarbe = 4                                                           'Gerbing 17.12.2017
                End If
                strAbleger = "Links"                                                                'Gerbing 25.09.2017
                StartPunktXLorum = cdLinks(AssPosition - 1).Left                                    'Gerbing 25.09.2017
                StartPunktYLorum = cdLinks(AssPosition - 1).Top                                     'Gerbing 25.09.2017
            End If
        Next I
    End If
    If PlayerTurn = Mitte Then
        For I = 1 To UBound(MitteCards)
            If MitteCards(I).cValue = 8 And MitteCards(I).cGespieltVon <> 9 Then                    '8=Ass
                blnAssGefunden = True
                AssPosition = I
                If MitteCards(AssPosition).cType = 1 Then
                    strFarbe = "Schell"
                    LorumAuflegeFarbe = 1                                                           'Gerbing 17.12.2017
                End If
                If MitteCards(AssPosition).cType = 2 Then
                    strFarbe = "Rot"
                    LorumAuflegeFarbe = 2                                                           'Gerbing 17.12.2017
                End If
                If MitteCards(AssPosition).cType = 3 Then
                    strFarbe = "Gr�n"
                    LorumAuflegeFarbe = 3  'Gerbing 17.12.2017
                End If
                If MitteCards(AssPosition).cType = 4 Then
                    strFarbe = "Eichel"
                    LorumAuflegeFarbe = 4                                                           'Gerbing 17.12.2017
                End If
                strAbleger = "Mitte"                                                                'Gerbing 25.09.2017
                StartPunktXLorum = cdMitte(AssPosition - 1).Left                                    'Gerbing 25.09.2017
                StartPunktYLorum = cdMitte(AssPosition - 1).Top                                     'Gerbing 25.09.2017
            End If
        Next I
    End If
    If PlayerTurn = Rechts Then
        For I = 1 To UBound(RechtsCards)
            If RechtsCards(I).cValue = 8 And RechtsCards(I).cGespieltVon <> 9 Then                    '8=Ass
                blnAssGefunden = True
                AssPosition = I
                If RechtsCards(AssPosition).cType = 1 Then
                    strFarbe = "Schell"
                    LorumAuflegeFarbe = 1                                                           'Gerbing 17.12.2017
                End If
                If RechtsCards(AssPosition).cType = 2 Then
                    strFarbe = "Rot"
                    LorumAuflegeFarbe = 2                                                           'Gerbing 17.12.2017
                End If
                If RechtsCards(AssPosition).cType = 3 Then
                    strFarbe = "Gr�n"
                    LorumAuflegeFarbe = 3                                                           'Gerbing 17.12.2017
                End If
                If RechtsCards(AssPosition).cType = 4 Then
                    strFarbe = "Eichel"
                    LorumAuflegeFarbe = 4  'Gerbing 17.12.2017
                End If
                strAbleger = "Rechts"                                                               'Gerbing 25.09.2017
                StartPunktXLorum = cdRechts(AssPosition - 1).Left                                   'Gerbing 25.09.2017
                StartPunktYLorum = cdRechts(AssPosition - 1).Top                                    'Gerbing 25.09.2017
            End If
        Next I
    End If
    If PlayerTurn = Spieler Then
        If LorumStartWert = 8 Then
            blnAssGefunden = True
            AssPosition = LorumStartPosition
            If SpielerCards(AssPosition).cType = 1 Then
                strFarbe = "Schell"
                LorumAuflegeFarbe = 1                                                               'Gerbing 17.12.2017
            End If
            If SpielerCards(AssPosition).cType = 2 Then
                strFarbe = "Rot"
                LorumAuflegeFarbe = 2                                                               'Gerbing 17.12.2017
            End If
            If SpielerCards(AssPosition).cType = 3 Then
                strFarbe = "Gr�n"
                LorumAuflegeFarbe = 3                                                               'Gerbing 17.12.2017
            End If
            If SpielerCards(AssPosition).cType = 4 Then
                strFarbe = "Eichel"
                LorumAuflegeFarbe = 4                                                               'Gerbing 17.12.2017
            End If
            strAbleger = "Spieler"                                                                  'Gerbing 25.09.2017
            StartPunktXLorum = cdSpieler(AssPosition - 1).Left                                      'Gerbing 25.09.2017
            StartPunktYLorum = cdSpieler(AssPosition - 1).Top                                       'Gerbing 25.09.2017
        Else
            For I = 1 To UBound(SpielerCards)
                If SpielerCards(I).cValue = 8 And SpielerCards(I).cGespieltVon <> 9 Then                    '8=Ass
                    blnAssGefunden = True
                    AssPosition = I
                    If SpielerCards(AssPosition).cType = 1 Then
                        strFarbe = "Schell"
                        LorumAuflegeFarbe = 1                                                       'Gerbing 17.12.2017
                    End If
                    If SpielerCards(AssPosition).cType = 2 Then
                        strFarbe = "Rot"
                        LorumAuflegeFarbe = 2                                                       'Gerbing 17.12.2017
                    End If
                    If SpielerCards(AssPosition).cType = 3 Then
                        strFarbe = "Gr�n"
                        LorumAuflegeFarbe = 3                                                       'Gerbing 17.12.2017
                    End If
                    If SpielerCards(AssPosition).cType = 4 Then
                        strFarbe = "Eichel"
                        LorumAuflegeFarbe = 4                                                       'Gerbing 17.12.2017
                    End If
                    strAbleger = "Spieler"                                                                  'Gerbing 25.09.2017
                    StartPunktXLorum = cdSpieler(AssPosition - 1).Left                                      'Gerbing 25.09.2017
                    StartPunktYLorum = cdSpieler(AssPosition - 1).Top                                       'Gerbing 25.09.2017
                End If
            Next I
        End If
    End If
    If blnAssGefunden = False Then
        LorumAssAblegen = 0
        Exit Function            'kein Ass gefunden
    End If
'------------------------------------------------------------------------------
    LorumStartWert = 8                                              'Gerbing 17.12.2017
    LorumAssAblegen = 1
    'hier das Ass auflegen
'    gstrMsg = strAbleger & " legt jetzt eine Karte ab" & vbNewLine                     'Gerbing 15.08.2017
'    gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
'    gstrMsg = gstrMsg & "Wert=" & strWert
'    frmMsgBox.Show 1
    '-------------------------------------------
    If strFarbe = "Schell" Then                                     '1=Schell
        ZielPunktLeftLorum = cdRoserlSchellAss(0).Left              'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlSchellAss(0).Top                'Gerbing 25.09.2017
        Call MoveKarteLorum(AssPosition - 1)                        'Gerbing 25.09.2017
        cdRoserlSchellAss(0).CardType = 1
        cdRoserlSchellAss(0).CardValue = 8
        cdRoserlSchellAss(0).FaceMode = NachOben
        cdRoserlSchellAss(0).Visible = True
    End If
    If strFarbe = "Rot" Then                                        '2=Rot
        ZielPunktLeftLorum = cdRoserlRotAss(0).Left                 'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlRotAss(0).Top                   'Gerbing 25.09.2017
        Call MoveKarteLorum(AssPosition - 1)                        'Gerbing 25.09.2017
        cdRoserlRotAss(0).CardType = 2
        cdRoserlRotAss(0).CardValue = 8
        cdRoserlRotAss(0).FaceMode = NachOben
        cdRoserlRotAss(0).Visible = True
    End If
    If strFarbe = "Gr�n" Then                                       '3=Gr�n
        ZielPunktLeftLorum = cdRoserlGr�nAss(0).Left                'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlGr�nAss(0).Top                  'Gerbing 25.09.2017
        Call MoveKarteLorum(AssPosition - 1)                        'Gerbing 25.09.2017
        cdRoserlGr�nAss(0).CardType = 3
        cdRoserlGr�nAss(0).CardValue = 8
        cdRoserlGr�nAss(0).FaceMode = NachOben
        cdRoserlGr�nAss(0).Visible = True
    End If
    If strFarbe = "Eichel" Then                                     '4=Eichel
        ZielPunktLeftLorum = cdRoserlEichelAss(0).Left              'Gerbing 25.09.2017
        ZielPunktTopLorum = cdRoserlEichelAss(0).Top                'Gerbing 25.09.2017
        Call MoveKarteLorum(AssPosition - 1)                        'Gerbing 25.09.2017
        cdRoserlEichelAss(0).CardType = 4
        cdRoserlEichelAss(0).CardValue = 8
        cdRoserlEichelAss(0).FaceMode = NachOben
        cdRoserlEichelAss(0).Visible = True
    End If
    '---------------------------------------------
    'Das abgelegte Ass aus ComputerCards entfernen
    'Call EntladenLorum(PlayerTurn, AssPosition)
    Call EntladenLorum(PlayerTurn)                                  'Gerbing 17.12.2017
End Function

Private Sub LorumEineKarteHinlegen(Index As Integer)
    'wenn sonst nichts mehr geht, dann muss der Computer eine Karte ablegen, ich nehme die erste die .cGespieltVon <> 9 ist
    'Wenn der Ausspieler nicht mehr weiter kann, kommt als n�chster Spieler der dran, der die letzte Karte einer vom
    'Ausspieler begonnenen Kartenfolge hinlegt, d.h. wenn die Kartenfolge nicht weitergeht, weil die ben�tigte Karte schon auf dem Tisch liegt.
    'Der Computer weis, wer die ben�tigte Karte hat
    'Zum Weiterspielen muss auf die Karten des vom Computer bestimmten Spielers geklickt werden
    '06.07.2017 Intelligenzverbesserung: Links Mite Rechts beginnt mit der ersten Karte(<>9) zu suchen ob es eine weitere Karte(<>9) der gleichen
    '   Farbe gibt, die nicht h�her als 3 Werte als die aktuelle ist oder nicht niedriger als 3 Werte als die aktuelle ist, die niedrigere
    '   von beiden wird zum Ablegen genommen

    Dim I As Integer
    Dim n As Integer
    Dim a As Integer
    Dim rc As Integer
    Dim strNextPlayer As String
    Dim strFarbe As String
    Dim strWert As String
    Dim blnN�chsteGefunden As Boolean
    Dim blnEsGibtZwei As Boolean

    a = 0
    LorumStartPosition = Index + 1
    blnEsGibtZwei = False
    If PlayerTurn = Spieler Then
        'beim Spieler braucht der Computer nicht zu kontrollieren ob eine Intelligenzverbesserung n�tig ist
        strAbleger = "Spieler"
        For I = 1 To 8
            If SpielerCards(LorumStartPosition).cGespieltVon <> 9 Then
                LorumAuflegeFarbe = SpielerCards(LorumStartPosition).cType                              'Farbe der geklickten Karte
                LorumStartWert = SpielerCards(LorumStartPosition).cValue                                'Wert der geklickten Karte
                Exit For
            End If
            LorumStartPosition = I
        Next I
    End If
    If PlayerTurn = Links Then
        strAbleger = "Links"
        For I = 1 To 8                                                                      'Gerbing 16.08.2017
            If LinksCards(I).cGespieltVon <> 9 Then
                a = a + 1
            End If
        Next I
        For I = 1 To 8                                                                      'Gerbing 06.07.2017
            If LinksCards(I).cGespieltVon <> 9 Then
                For n = 1 To 8
                    If I <> n Then
                        If LinksCards(n).cGespieltVon <> 9 Then
                            If LinksCards(I).cType = LinksCards(n).cType Then
                                If LinksCards(I).cValue <= LinksCards(n).cValue + 3 Then
                                    blnEsGibtZwei = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next n
                If blnEsGibtZwei = True Then Exit For
            End If
        Next I
        If blnEsGibtZwei = True Then
            If LinksCards(I).cValue <= LinksCards(n).cValue Then
                LorumStartPosition = I
            Else
                LorumStartPosition = n
            End If
            LorumAuflegeFarbe = LinksCards(LorumStartPosition).cType                              'Farbe der geklickten Karte
            LorumStartWert = LinksCards(LorumStartPosition).cValue                                'Wert der geklickten Karte
        Else
            For I = 1 To 8
                If LinksCards(I).cGespieltVon <> 9 Then
                    LorumAuflegeFarbe = LinksCards(I).cType                              'Farbe der geklickten Karte
                    LorumStartWert = LinksCards(I).cValue                                'Wert der geklickten Karte
                    Exit For
                End If
                LorumStartPosition = I
            Next I
        End If
    End If
    If PlayerTurn = Mitte Then
        strAbleger = "Mitte"
        For I = 1 To 8                                                                      'Gerbing 16.08.2017
            If MitteCards(I).cGespieltVon <> 9 Then
                a = a + 1
            End If
        Next I
        For I = 1 To 8                                                                      'Gerbing 06.07.2017
            If MitteCards(I).cGespieltVon <> 9 Then
                a = a + 1                                                                   'Gerbing 16.08.2017
                For n = 1 To 8
                    If I <> n Then
                        If MitteCards(n).cGespieltVon <> 9 Then
                            If MitteCards(I).cType = MitteCards(n).cType Then
                                If MitteCards(I).cValue <= MitteCards(n).cValue + 3 Then
                                    blnEsGibtZwei = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next n
                If blnEsGibtZwei = True Then Exit For
            End If
        Next I
        If blnEsGibtZwei = True Then
            If MitteCards(I).cValue <= MitteCards(n).cValue Then
                LorumStartPosition = I
            Else
                LorumStartPosition = n
            End If
            LorumAuflegeFarbe = MitteCards(LorumStartPosition).cType                              'Farbe der geklickten Karte
            LorumStartWert = MitteCards(LorumStartPosition).cValue                                'Wert der geklickten Karte
        Else
            For I = 1 To 8
                If MitteCards(I).cGespieltVon <> 9 Then
                    LorumAuflegeFarbe = MitteCards(I).cType                              'Farbe der geklickten Karte
                    LorumStartWert = MitteCards(I).cValue                                'Wert der geklickten Karte
                    Exit For
                End If
                LorumStartPosition = I
            Next I
        End If
    End If
    If PlayerTurn = Rechts Then
        strAbleger = "Rechts"
        For I = 1 To 8                                                                      'Gerbing 16.08.2017
            If RechtsCards(I).cGespieltVon <> 9 Then
                a = a + 1
            End If
        Next I
        For I = 1 To 8                                                                      'Gerbing 06.07.2017
            If RechtsCards(I).cGespieltVon <> 9 Then
                a = a + 1                                                                   'Gerbing 16.08.2017
                For n = 1 To 8
                    If I <> n Then
                        If RechtsCards(n).cGespieltVon <> 9 Then
                            If RechtsCards(I).cType = RechtsCards(n).cType Then
                                If RechtsCards(I).cValue <= RechtsCards(n).cValue + 3 Then
                                    blnEsGibtZwei = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next n
                If blnEsGibtZwei = True Then Exit For
            End If
        Next I
        If blnEsGibtZwei = True Then
            If RechtsCards(I).cValue <= RechtsCards(n).cValue Then
                LorumStartPosition = I
            Else
                LorumStartPosition = n
            End If
            LorumAuflegeFarbe = RechtsCards(LorumStartPosition).cType                              'Farbe der geklickten Karte
            LorumStartWert = RechtsCards(LorumStartPosition).cValue                                'Wert der geklickten Karte
        Else
            For I = 1 To 8
                If RechtsCards(I).cGespieltVon <> 9 Then
                    LorumAuflegeFarbe = RechtsCards(I).cType                              'Farbe der geklickten Karte
                    LorumStartWert = RechtsCards(I).cValue                                'Wert der geklickten Karte
                    Exit For
                End If
                LorumStartPosition = I
            Next I
        End If
    End If
    If LorumAuflegeFarbe = 1 Then strFarbe = "Schell"
    If LorumAuflegeFarbe = 2 Then strFarbe = "Rot"
    If LorumAuflegeFarbe = 3 Then strFarbe = "Gr�n"
    If LorumAuflegeFarbe = 4 Then strFarbe = "Eichel"
    If LorumStartWert = 1 Then strWert = "Sieben"
    If LorumStartWert = 2 Then strWert = "Acht"
    If LorumStartWert = 3 Then strWert = "Neun"
    If LorumStartWert = 4 Then strWert = "Zehn"
    If LorumStartWert = 5 Then strWert = "Unter"
    If LorumStartWert = 6 Then strWert = "Ober"
    If LorumStartWert = 7 Then strWert = "K�nig"
    If LorumStartWert = 8 Then strWert = "Ass"
'    If a = 1 Then
'        gstrMsg = strAbleger & " legt jetzt seine letzte Karte ab"                      'Gerbing 16.08.2017
'    Else
'        gstrMsg = strAbleger & " legt jetzt eine Karte ab" & vbNewLine
'        gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
'        gstrMsg = gstrMsg & "Wert=" & strWert
'    End If
'    frmMsgBox.Show 1
    '----------------------------------------------------------------------------------
    'F�r MoveKarteLorum den Startpunkt festlegen                                        'Gerbing 25.09.2017
    If strAbleger = "Links" Then                                                        'Gerbing 25.09.2017
        StartPunktXLorum = cdLinks(Index).Left                                          'Gerbing 25.09.2017
        StartPunktYLorum = cdLinks(Index).Top                                           'Gerbing 25.09.2017
    ElseIf strAbleger = "Mitte" Then
        StartPunktXLorum = cdMitte(Index).Left                                          'Gerbing 25.09.2017
        StartPunktYLorum = cdMitte(Index).Top                                           'Gerbing 25.09.2017
    ElseIf strAbleger = "Rechts" Then
        StartPunktXLorum = cdRechts(Index).Left                                          'Gerbing 25.09.2017
        StartPunktYLorum = cdRechts(Index).Top                                           'Gerbing 25.09.2017
    ElseIf strAbleger = "Spieler" Then
        StartPunktXLorum = cdSpieler(Index).Left                                          'Gerbing 25.09.2017
        StartPunktYLorum = cdSpieler(Index).Top                                           'Gerbing 25.09.2017
    End If
    '----------------------------------------------------------------------------------
    Do
        'Ich kann die Funktion LorumKartenFolge dazu benutzen jeweils eine Anlegekarte abzulegen dazu muss der Parameter
        'blnLoopBenutzen = False sein
        'LorumAuflegeFarbe LorumStartWert LorumStartPosition werden auch in der Funktion LorumKartenFolge benutzt
        rc = LorumKartenFolge(False)       'false=Loop nicht benutzen
        'rc=0=dies war die letzte Karte der Kartenfolge
        'rc=1=es gibt eine weitere Karte in der Kartenfolge
        If blnLorumSpielende = True Then Exit Sub
        If rc = 0 Then
            Exit Do
        End If
    Loop
    'wenn die erste und letzte Karte einer 4er-Folge vom gleichen Mitspieler hingelegt werden, kommt dieser Mitspieler weiter dran
    If blnErsteKommtVonLinks = True And PlayerTurn = Links And LorumAufgelegtZ�hler = 0 Then    'Bei =0 war es die letzte einer Folge
        LorumAufgelegtZ�hler = 1
        blnErsteKommtVonLinks = False
        GoTo N�chsteGefunden
    End If
    If blnErsteKommtVonMitte = True And PlayerTurn = Mitte And LorumAufgelegtZ�hler = 0 Then
        LorumAufgelegtZ�hler = 1
        blnErsteKommtVonMitte = False
        GoTo N�chsteGefunden
    End If
    If blnErsteKommtVonRechts = True And PlayerTurn = Rechts And LorumAufgelegtZ�hler = 0 Then
        LorumAufgelegtZ�hler = 1
        blnErsteKommtVonRechts = False
        GoTo N�chsteGefunden
    End If
    'Wenn Spieler selbst die erste Karte eine 4er-Folge hingelegt hat, braucht nicht gesucht zu werden wem die n�chste Karte geh�rt
    If blnErsteKommtVonSpieler = True Then
        LorumAufgelegtZ�hler = 1
        blnN�chsteGefunden = False
    Else
        '---------------------------------------------------------------------------------
        'Wem geh�rt die n�chste Karte
        LorumStartWert = LorumStartWert + 1
        blnN�chsteGefunden = False
        If LorumStartWert = 9 Then                                          'Gerbing 05.10.2017
            GoTo N�chsteGefunden
        End If
        For n = 1 To UBound(LinksCards)
            If LinksCards(n).cValue = LorumStartWert And LinksCards(n).cType = LorumAuflegeFarbe And LinksCards(n).cGespieltVon <> 9 Then
                LorumStartPosition = n                                      '30.06.2017 LorumStartPosition war falsch wenn UBound(MitteCards) falsch war
                PlayerTurnRoserl = 1                                        '1=Links
                blnN�chsteGefunden = True
                blnN�chsteGefunden = True
                If LinksCards(n).cType = 1 Then strFarbe = "Schell"
                If LinksCards(n).cType = 2 Then strFarbe = "Rot"
                If LinksCards(n).cType = 3 Then strFarbe = "Gr�n"
                If LinksCards(n).cType = 4 Then strFarbe = "Eichel"
                If LorumStartWert = 1 Then strWert = "Sieben"
                If LorumStartWert = 2 Then strWert = "Acht"
                If LorumStartWert = 3 Then strWert = "Neun"
                If LorumStartWert = 4 Then strWert = "Zehn"
                If LorumStartWert = 5 Then strWert = "Unter"
                If LorumStartWert = 6 Then strWert = "Ober"
                If LorumStartWert = 7 Then strWert = "K�nig"
                If LorumStartWert = 8 Then strWert = "Ass"
                gstrMsg = "Der Computer legt jetzt eine Karte von Links ab" & vbNewLine
                gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
                gstrMsg = gstrMsg & "Wert=" & strWert
                frmMsgBox.Show 1
                GoTo N�chsteGefunden
            End If
        Next n
        For n = 1 To UBound(MitteCards)
            If MitteCards(n).cValue = LorumStartWert And MitteCards(n).cType = LorumAuflegeFarbe And MitteCards(n).cGespieltVon <> 9 Then
                LorumStartPosition = n                                      '30.06.2017 LorumStartPosition war falsch wenn UBound(MitteCards) falsch war
                PlayerTurnRoserl = 2                                        '2=mitte
                blnN�chsteGefunden = True
                If MitteCards(n).cType = 1 Then strFarbe = "Schell"
                If MitteCards(n).cType = 2 Then strFarbe = "Rot"
                If MitteCards(n).cType = 3 Then strFarbe = "Gr�n"
                If MitteCards(n).cType = 4 Then strFarbe = "Eichel"
                If LorumStartWert = 1 Then strWert = "Sieben"
                If LorumStartWert = 2 Then strWert = "Acht"
                If LorumStartWert = 3 Then strWert = "Neun"
                If LorumStartWert = 4 Then strWert = "Zehn"
                If LorumStartWert = 5 Then strWert = "Unter"
                If LorumStartWert = 6 Then strWert = "Ober"
                If LorumStartWert = 7 Then strWert = "K�nig"
                If LorumStartWert = 8 Then strWert = "Ass"
                gstrMsg = "Der Computer legt jetzt eine Karte von Mitte ab" & vbNewLine
                gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
                gstrMsg = gstrMsg & "Wert=" & strWert
                frmMsgBox.Show 1
                GoTo N�chsteGefunden
            End If
        Next n
        For n = 1 To UBound(RechtsCards)
            If RechtsCards(n).cValue = LorumStartWert And RechtsCards(n).cType = LorumAuflegeFarbe And RechtsCards(n).cGespieltVon <> 9 Then
                LorumStartPosition = n                                      '30.06.2017 LorumStartPosition war falsch wenn UBound(RechtsCards) falsch war
                PlayerTurnRoserl = 3                                        '3=mitte
                blnN�chsteGefunden = True
                If RechtsCards(n).cType = 1 Then strFarbe = "Schell"
                If RechtsCards(n).cType = 2 Then strFarbe = "Rot"
                If RechtsCards(n).cType = 3 Then strFarbe = "Gr�n"
                If RechtsCards(n).cType = 4 Then strFarbe = "Eichel"
                If LorumStartWert = 1 Then strWert = "Sieben"
                If LorumStartWert = 2 Then strWert = "Acht"
                If LorumStartWert = 3 Then strWert = "Neun"
                If LorumStartWert = 4 Then strWert = "Zehn"
                If LorumStartWert = 5 Then strWert = "Unter"
                If LorumStartWert = 6 Then strWert = "Ober"
                If LorumStartWert = 7 Then strWert = "K�nig"
                If LorumStartWert = 8 Then strWert = "Ass"
                gstrMsg = "Der Computer legt jetzt eine Karte von Rechts ab" & vbNewLine
                gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
                gstrMsg = gstrMsg & "Wert=" & strWert
                frmMsgBox.Show 1
                GoTo N�chsteGefunden
            End If
        Next n
        '---------------------------------------------------
        For n = 1 To UBound(SpielerCards)
            On Error Resume Next                            'Gerbing 29.06.2017
            SpielerCards(n).cValue = cdSpieler(n - 1).CardValue
            SpielerCards(n).cType = cdSpieler(n - 1).CardType
            On Error GoTo 0
            If SpielerCards(n).cValue = LorumStartWert And SpielerCards(n).cType = LorumAuflegeFarbe And SpielerCards(n).cGespieltVon <> 9 Then
                LorumStartPosition = n                                      '30.06.2017 LorumStartPosition war falsch wenn UBound(SpielerCards) falsch war
                PlayerTurnRoserl = 0                                        '0=Spieler
                blnN�chsteGefunden = True
                If SpielerCards(n).cType = 1 Then strFarbe = "Schell"
                If SpielerCards(n).cType = 2 Then strFarbe = "Rot"
                If SpielerCards(n).cType = 3 Then strFarbe = "Gr�n"
                If SpielerCards(n).cType = 4 Then strFarbe = "Eichel"
                If LorumStartWert = 1 Then strWert = "Sieben"
                If LorumStartWert = 2 Then strWert = "Acht"
                If LorumStartWert = 3 Then strWert = "Neun"
                If LorumStartWert = 4 Then strWert = "Zehn"
                If LorumStartWert = 5 Then strWert = "Unter"
                If LorumStartWert = 6 Then strWert = "Ober"
                If LorumStartWert = 7 Then strWert = "K�nig"
                If LorumStartWert = 8 Then strWert = "Ass"
                gstrMsg = "Der Computer legt jetzt eine Karte von Dir ab" & vbNewLine
                gstrMsg = gstrMsg & "Farbe=" & strFarbe & vbNewLine
                gstrMsg = gstrMsg & "Wert=" & strWert
                frmMsgBox.Show 1
                Exit For
            End If
        Next n
N�chsteGefunden:
    End If
    If PlayerTurnRoserl = 1 Then
        strNextPlayer = "Links"
        lblLinks.Font.Bold = True                           'Gerbing 15.08.2017
        lblLinks.ForeColor = FarbeAusspieler                           'schwarz
        lblSpieler.ForeColor = &HFFFFFF                     'wei�
        lblMitte.ForeColor = &HFFFFFF
        lblRechts.ForeColor = &HFFFFFF
        lblMitte.Font.Bold = False
        lblRechts.Font.Bold = False
        lblSpieler.Font.Bold = False
    End If
    If PlayerTurnRoserl = 2 Then
        strNextPlayer = "Mitte"
        lblMitte.Font.Bold = True                           'Gerbing 15.08.2017
        lblMitte.ForeColor = FarbeAusspieler                           'schwarz
        lblLinks.ForeColor = &HFFFFFF                       'wei�
        lblSpieler.ForeColor = &HFFFFFF
        lblRechts.ForeColor = &HFFFFFF
        lblLinks.Font.Bold = False
        lblRechts.Font.Bold = False
        lblSpieler.Font.Bold = False
    End If
    If PlayerTurnRoserl = 3 Then
        strNextPlayer = "Rechts"
        lblRechts.Font.Bold = True                          'Gerbing 15.08.2017
        lblRechts.ForeColor = FarbeAusspieler                          'schwarz
        lblLinks.ForeColor = &HFFFFFF                       'wei�
        lblMitte.ForeColor = &HFFFFFF
        lblSpieler.ForeColor = &HFFFFFF
        lblLinks.Font.Bold = False
        lblMitte.Font.Bold = False
        lblSpieler.Font.Bold = False
    End If
    If PlayerTurnRoserl = 0 Then
        strNextPlayer = "Spieler"
        lblSpieler.Font.Bold = True                         'Gerbing 15.08.2017
        lblSpieler.ForeColor = FarbeAusspieler                         'schwarz
        lblLinks.ForeColor = &HFFFFFF                       'wei�
        lblMitte.ForeColor = &HFFFFFF
        lblRechts.ForeColor = &HFFFFFF
        lblLinks.Font.Bold = False
        lblMitte.Font.Bold = False
        lblRechts.Font.Bold = False
    End If
    gstrMsg = "Der n�chste Ausspieler ist " & strNextPlayer & vbNewLine
    'gstrMsg = gstrMsg & "weil er die n�chste zum Bilden einer Folge ben�tigte Karte hat" & vbNewLine
    'gstrMsg = gstrMsg & "oder weil er die letzte Karte einer aufsteigenden Folge hingelegt hat." & vbNewLine
    'gstrMsg = gstrMsg & "Du musst auf die Karten des Ausspielers Klicken" & vbNewLine  'Gerbing 04.10.2017
    frmMsgBox.Show 1
    '--------------------------------------------------------Gerbing 04.10.2017
    blnloruml�uft = False                                   'Gerbing 14.11.2017
    If strNextPlayer = "Links" Then
        Call cdLinks_Click(0)
    ElseIf strNextPlayer = "Mitte" Then
        Call cdMitte_Click(0)
    ElseIf strNextPlayer = "Rechts" Then
        Call cdRechts_Click(0)
    End If
    'blnloruml�uft = False                                   'Gerbing 14.11.2017
End Sub

Private Sub mnuBeenden_Click()
    Call SpeichernSpielstand
    End
End Sub

Private Sub mnuBestwerte_Click()
    frmBestwerteZeigen.Show 1
End Sub

Private Sub mnuFontSize_Click()
    frmFontSize.Show 1
End Sub

Private Sub mnuLetztenStich_Click()
    If AKIndex >= 4 Then
        frmLetztenStichZeigen.Show 1
    End If
End Sub

Private Sub mnuLorumkartenAufgedeckt_Click()

End Sub

Private Sub mnuSpielbogen_Click()
    frmSpielBogen.Show
End Sub

Private Sub mnuSpielregeln_Click()
    Dim strHilfefile As String
    Dim strDateinamenErweiterung As String
    Dim retVal As Long
    Dim ErrorText As String
    
    strHilfefile = App.Path & "\LORUM am Computer.pdf"
    retVal = ShellExecute(Me.hWnd, "open", strHilfefile, vbNull, vbNull, 1)
    If retVal <= 32 Then
        strDateinamenErweiterung = "pdf"
        ErrorText = GetShellError(retVal)
        gstrMsg = "Errortext=" & ErrorText & vbNewLine
        gstrMsg = gstrMsg & "Errornr=" & retVal & vbNewLine & vbNewLine
        
        gstrMsg = gstrMsg & strHilfefile & vbNewLine
        gstrMsg = gstrMsg & "Diese Datei kann nicht ge�ffnet werden." & vbNewLine & vbNewLine
        
        gstrMsg = gstrMsg & "Entweder die Datei existiert nicht," & vbNewLine & vbNewLine
        
        gstrMsg = gstrMsg & "oder es ist keine Anwendung mit der" & vbNewLine
        gstrMsg = gstrMsg & "Dateinamen-Erweiterung(Datei-Typ) " & strDateinamenErweiterung & " verkn�pft." & vbNewLine
        gstrMsg = gstrMsg & "W�hle selbst eine geignete Anwendung, zB mittels Windows-Explorer" & vbNewLine
        gstrMsg = gstrMsg & "Rechtklicken auf den Dateiname -> �ffnen mit... -> Programm ausw�hlen"
        frmMsgBox.Show 1
    End If
End Sub

Private Sub mnuVersion_Click()
    'Versions-Informationen ermitteln                           'Gerbing 12.09.2017
    gstrMsg = "Version " & GetRoserlExeVersion                  'Gerbing 12.09.2017
    frmMsgBox.Show 1
End Sub

Private Sub SchlafTimer_Timer()
    SchlafTimer.Enabled = False
    If App.LogMode = 0 Then                             'App.LogMode = 0 'ich bin in der IDE
        Exit Sub
    End If
    gstrMsg = "Schlaf nicht, Du bist dran"
    On Error Resume Next                                'Gerbing 22.09.2017
    frmMsgBox.Show 1
End Sub

Private Sub PunkteZ�hlen()
    'Auswertung der 32 Karten von AusgespielteKarten(i).cEingesammeltVon
    Dim I As Integer
    Dim strSpielart As String
    
    LinksPunkte = 0                                     'Gerbing 26.07.2020
    MittePunkte = 0                                     'Gerbing 26.07.2020
    RechtsPunkte = 0                                    'Gerbing 26.07.2020
    SpielerPunkte = 0                                   'Gerbing 26.07.2020
    
    lblSpieler.Font.Bold = False                        'Gerbing 15.08.2017
    lblSpieler.ForeColor = &HFFFFFF                     'wei�
    lblLinks.Font.Bold = False
    lblLinks.ForeColor = &HFFFFFF
    lblMitte.Font.Bold = False
    lblMitte.ForeColor = &HFFFFFF
    lblRechts.Font.Bold = False
    lblRechts.ForeColor = &HFFFFFF
    If Spielart = Lorum Then
        'Jeder Mitspieler z�hlt seine Karten, das sind die Strafpunkte
        If LorumGewinner = Links Then
            LinksPunkte = 0
        Else
            LinksPunkte = 0
            For I = 0 To cdLinks.Count - 1
                On Error Resume Next
                If cdLinks(I).Visible = True Then
                    LinksPunkte = LinksPunkte + 1
                End If
                On Error GoTo 0
            Next I
        End If
        '----------------------------
        If LorumGewinner = Mitte Then
            MittePunkte = 0
        Else
            MittePunkte = 0
            For I = 0 To cdMitte.Count - 1
                On Error Resume Next
                If cdMitte(I).Visible = True Then
                    MittePunkte = MittePunkte + 1
                End If
                On Error GoTo 0
            Next I
        End If
        '-----------------------------
        If LorumGewinner = Rechts Then
            RechtsPunkte = 0
        Else
            RechtsPunkte = 0
            For I = 0 To cdRechts.Count - 1
                On Error Resume Next
                If cdRechts(I).Visible = True Then
                    RechtsPunkte = RechtsPunkte + 1
                End If
                On Error GoTo 0
            Next I
        End If
        '------------------------------
        If LorumGewinner = Spieler Then
            SpielerPunkte = 0
        Else
            SpielerPunkte = 0
            For I = 0 To cdSpieler.Count - 1
                On Error Resume Next
                If cdSpieler(I).Visible = True Then
                    SpielerPunkte = SpielerPunkte + 1
                End If
                On Error GoTo 0
            Next I
        End If
    End If
    '----------------------------------
    If Spielart <> Lorum Then
        For I = 1 To AKIndex
            If Spielart = Rot Then
                'jede rote Karte kostet einen Punkt
                If AusgespielteKarten(I).cType = 2 Then                             '2=Rot
                    If AusgespielteKarten(I).cEingesammeltVon = Links Then
                        LinksPunkte = LinksPunkte + 1
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Mitte Then
                        MittePunkte = MittePunkte + 1
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Rechts Then
                        RechtsPunkte = RechtsPunkte + 1
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Spieler Then
                        SpielerPunkte = SpielerPunkte + 1
                    End If
                End If
            End If
            If Spielart = Ober Then
                'jeder Ober kostet zwei Punkte
                If AusgespielteKarten(I).cValue = 6 Then                            '6=Ober
                    If AusgespielteKarten(I).cEingesammeltVon = Links Then
                        LinksPunkte = LinksPunkte + 2
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Mitte Then
                        MittePunkte = MittePunkte + 2
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Rechts Then
                        RechtsPunkte = RechtsPunkte + 2
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Spieler Then
                        SpielerPunkte = SpielerPunkte + 2
                    End If
                End If
            End If
            If Spielart = Stiche Then
                'jeder Stich kostet einen Punkt
                If AusgespielteKarten(I).cEingesammeltVon = Links Then
                    LinksPunkte = LinksPunkte + 1
                ElseIf AusgespielteKarten(I).cEingesammeltVon = Mitte Then
                    MittePunkte = MittePunkte + 1
                ElseIf AusgespielteKarten(I).cEingesammeltVon = Rechts Then
                    RechtsPunkte = RechtsPunkte + 1
                ElseIf AusgespielteKarten(I).cEingesammeltVon = Spieler Then
                    SpielerPunkte = SpielerPunkte + 1
                End If
            End If
            If Spielart = ErsterUndLetzterstich Then
                'jeweils der erste und der letzte Stich kostet 4 Punkte
                If (I >= 1 And I <= 4) Or (I >= 29 And I <= 32) Then
                    If AusgespielteKarten(I).cEingesammeltVon = Links Then
                        LinksPunkte = LinksPunkte + 4
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Mitte Then
                        MittePunkte = MittePunkte + 4
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Rechts Then
                        RechtsPunkte = RechtsPunkte + 4
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Spieler Then
                        SpielerPunkte = SpielerPunkte + 4
                    End If
                End If
            End If
            If Spielart = RoterK�nig Then
                'Der RoteK�nig kostet 8 Punkte
                If AusgespielteKarten(I).cValue = 7 And AusgespielteKarten(I).cType = 2 Then                '8=K�nig,2=Rot
                    If AusgespielteKarten(I).cEingesammeltVon = Links Then
                        LinksPunkte = LinksPunkte + 8
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Mitte Then
                        MittePunkte = MittePunkte + 8
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Rechts Then
                        RechtsPunkte = RechtsPunkte + 8
                    ElseIf AusgespielteKarten(I).cEingesammeltVon = Spieler Then
                        SpielerPunkte = SpielerPunkte + 8
                    End If
                End If
            End If
        Next I
    End If
    '-----------------------------------------------------------------------------------------------
    If Spielart = Rot Then
        strSpielart = "Rot"
    ElseIf Spielart = Ober Then
        strSpielart = "Ober"
    ElseIf Spielart = Stiche Then
        strSpielart = "Stiche"
        LinksPunkte = LinksPunkte \ 4
        MittePunkte = MittePunkte \ 4
        RechtsPunkte = RechtsPunkte \ 4
        SpielerPunkte = SpielerPunkte \ 4
    ElseIf Spielart = ErsterUndLetzterstich Then
        strSpielart = "Erster und letzter Stich"
        LinksPunkte = LinksPunkte \ 4
        MittePunkte = MittePunkte \ 4
        RechtsPunkte = RechtsPunkte \ 4
        SpielerPunkte = SpielerPunkte \ 4
    ElseIf Spielart = RoterK�nig Then
        strSpielart = "Roter K�nig"
    ElseIf Spielart = Lorum Then
        strSpielart = "Lorum"
    End If
    gstrMsg = "Spielart=" & strSpielart & vbNewLine & vbNewLine
    gstrMsg = gstrMsg & "Links=" & LinksPunkte & " Punkte" & vbNewLine
    gstrMsg = gstrMsg & "Mitte=" & MittePunkte & " Punkte" & vbNewLine
    gstrMsg = gstrMsg & "Rechts=" & RechtsPunkte & " Punkte" & vbNewLine
    gstrMsg = gstrMsg & "Spieler=" & SpielerPunkte & " Punkte" & vbNewLine
    '-------------------------------------------------------------
    'Gerbing 26.07.2020
    'anstelle frmMsgBox.Show 1 soll eine neue Form frmSpielEnde benutzt werden, wo der Nutzter 'Spiel wiederholen' ausw�hlen kann
    'frmMsgBox.Show 1
    frmSpielEnde.Show 1                                                 'Gerbing 26.07.2020
    If blnSpielWiederholen = True Then                                  'Gerbing 26.07.2020
        Exit Sub                                                        'Gerbing 26.07.2020
    End If                                                              'Gerbing 26.07.2020
    '-------------------------------------------------------------
    'Punkte in den Spiel-Bogen eintragen
    If Spielart = Rot Then
        If Rundez�hler = 1 Then
            frmSpielBogen.lblLinksRot_1 = LinksPunkte
            frmSpielBogen.lblMitteRot_1 = MittePunkte
            frmSpielBogen.lblRechtsRot_1 = RechtsPunkte
            frmSpielBogen.lblSpielerRot_1 = SpielerPunkte
        ElseIf Rundez�hler = 2 Then
            frmSpielBogen.lblLinksRot_2 = LinksPunkte
            frmSpielBogen.lblMitteRot_2 = MittePunkte
            frmSpielBogen.lblRechtsRot_2 = RechtsPunkte
            frmSpielBogen.lblSpielerRot_2 = SpielerPunkte
        ElseIf Rundez�hler = 3 Then
            frmSpielBogen.lblLinksRot_3 = LinksPunkte
            frmSpielBogen.lblMitteRot_3 = MittePunkte
            frmSpielBogen.lblRechtsRot_3 = RechtsPunkte
            frmSpielBogen.lblSpielerRot_3 = SpielerPunkte
        ElseIf Rundez�hler = 4 Then
            frmSpielBogen.lblLinksRot_4 = LinksPunkte
            frmSpielBogen.lblMitteRot_4 = MittePunkte
            frmSpielBogen.lblRechtsRot_4 = RechtsPunkte
            frmSpielBogen.lblSpielerRot_4 = SpielerPunkte
        End If
    ElseIf Spielart = Ober Then
        If Rundez�hler = 1 Then
            frmSpielBogen.lblLinksOber_1 = LinksPunkte
            frmSpielBogen.lblMItteOber_1 = MittePunkte
            frmSpielBogen.lblRechtsOber_1 = RechtsPunkte
            frmSpielBogen.lblSpielerOber_1 = SpielerPunkte
        ElseIf Rundez�hler = 2 Then
            frmSpielBogen.lblLinksOber_2 = LinksPunkte
            frmSpielBogen.lblMItteOber_2 = MittePunkte
            frmSpielBogen.lblRechtsOber_2 = RechtsPunkte
            frmSpielBogen.lblSpielerOber_2 = SpielerPunkte
        ElseIf Rundez�hler = 3 Then
            frmSpielBogen.lblLinksOber_3 = LinksPunkte
            frmSpielBogen.lblMItteOber_3 = MittePunkte
            frmSpielBogen.lblRechtsOber_3 = RechtsPunkte
            frmSpielBogen.lblSpielerOber_3 = SpielerPunkte
        ElseIf Rundez�hler = 4 Then
            frmSpielBogen.lblLinksOber_4 = LinksPunkte
            frmSpielBogen.lblMItteOber_4 = MittePunkte
            frmSpielBogen.lblRechtsOber_4 = RechtsPunkte
            frmSpielBogen.lblSpielerOber_4 = SpielerPunkte
        End If
    ElseIf Spielart = Stiche Then
        If Rundez�hler = 1 Then
            frmSpielBogen.lblLinksStiche_1 = LinksPunkte
            frmSpielBogen.lblMitteStiche_1 = MittePunkte
            frmSpielBogen.lblRechtsStiche_1 = RechtsPunkte
            frmSpielBogen.lblSpielerStiche_1 = SpielerPunkte
        ElseIf Rundez�hler = 2 Then
            frmSpielBogen.lblLinksStiche_2 = LinksPunkte
            frmSpielBogen.lblMitteStiche_2 = MittePunkte
            frmSpielBogen.lblRechtsStiche_2 = RechtsPunkte
            frmSpielBogen.lblSpielerStiche_2 = SpielerPunkte
        ElseIf Rundez�hler = 3 Then
            frmSpielBogen.lblLinksStiche_3 = LinksPunkte
            frmSpielBogen.lblMitteStiche_3 = MittePunkte
            frmSpielBogen.lblRechtsStiche_3 = RechtsPunkte
            frmSpielBogen.lblSpielerStiche_3 = SpielerPunkte
        ElseIf Rundez�hler = 4 Then
            frmSpielBogen.lblLinksStiche_4 = LinksPunkte
            frmSpielBogen.lblMitteStiche_4 = MittePunkte
            frmSpielBogen.lblRechtsStiche_4 = RechtsPunkte
            frmSpielBogen.lblSpielerStiche_4 = SpielerPunkte
        End If
    ElseIf Spielart = ErsterUndLetzterstich Then
        If Rundez�hler = 1 Then
            frmSpielBogen.lblLinksEuL_1 = LinksPunkte
            frmSpielBogen.lblMitteEuL_1 = MittePunkte
            frmSpielBogen.lblRechtsEuL_1 = RechtsPunkte
            frmSpielBogen.lblSpielerEuL_1 = SpielerPunkte
        ElseIf Rundez�hler = 2 Then
            frmSpielBogen.lblLinksEuL_2 = LinksPunkte
            frmSpielBogen.lblMitteEuL_2 = MittePunkte
            frmSpielBogen.lblRechtsEuL_2 = RechtsPunkte
            frmSpielBogen.lblSpielerEuL_2 = SpielerPunkte
        ElseIf Rundez�hler = 3 Then
            frmSpielBogen.lblLinksEuL_3 = LinksPunkte
            frmSpielBogen.lblMitteEuL_3 = MittePunkte
            frmSpielBogen.lblRechtsEuL_3 = RechtsPunkte
            frmSpielBogen.lblSpielerEuL_3 = SpielerPunkte
        ElseIf Rundez�hler = 4 Then
            frmSpielBogen.lblLinksEuL_4 = LinksPunkte
            frmSpielBogen.lblMitteEuL_4 = MittePunkte
            frmSpielBogen.lblRechtsEuL_4 = RechtsPunkte
            frmSpielBogen.lblSpielerEuL_4 = SpielerPunkte
        End If
    ElseIf Spielart = RoterK�nig Then
        If Rundez�hler = 1 Then
            frmSpielBogen.lblLinksRK_1 = LinksPunkte
            frmSpielBogen.lblMitteRK_1 = MittePunkte
            frmSpielBogen.lblRechtsRK_1 = RechtsPunkte
            frmSpielBogen.lblSpielerRK_1 = SpielerPunkte
        ElseIf Rundez�hler = 2 Then
            frmSpielBogen.lblLinksRK_2 = LinksPunkte
            frmSpielBogen.lblMitteRK_2 = MittePunkte
            frmSpielBogen.lblRechtsRK_2 = RechtsPunkte
            frmSpielBogen.lblSpielerRK_2 = SpielerPunkte
        ElseIf Rundez�hler = 3 Then
            frmSpielBogen.lblLinksRK_3 = LinksPunkte
            frmSpielBogen.lblMitteRK_3 = MittePunkte
            frmSpielBogen.lblRechtsRK_3 = RechtsPunkte
            frmSpielBogen.lblSpielerRK_3 = SpielerPunkte
        ElseIf Rundez�hler = 4 Then
            frmSpielBogen.lblLinksRK_4 = LinksPunkte
            frmSpielBogen.lblMitteRK_4 = MittePunkte
            frmSpielBogen.lblRechtsRK_4 = RechtsPunkte
            frmSpielBogen.lblSpielerRK_4 = SpielerPunkte
        End If
    ElseIf Spielart = Lorum Then
        If Rundez�hler = 1 Then
            frmSpielBogen.lblLinksLorum_1 = LinksPunkte
            frmSpielBogen.lblMitteLorum_1 = MittePunkte
            frmSpielBogen.lblRechtsLorum_1 = RechtsPunkte
            frmSpielBogen.lblSpielerLorum_1 = SpielerPunkte
            frmSpielBogen.lblLinksSumme_1 = CInt(frmSpielBogen.lblLinksRot_1) + CInt(frmSpielBogen.lblLinksOber_1) + CInt(frmSpielBogen.lblLinksStiche_1) + CInt(frmSpielBogen.lblLinksEuL_1) + CInt(frmSpielBogen.lblLinksRK_1) + CInt(frmSpielBogen.lblLinksLorum_1)
            frmSpielBogen.lblMitteSumme_1 = CInt(frmSpielBogen.lblMitteRot_1) + CInt(frmSpielBogen.lblMItteOber_1) + CInt(frmSpielBogen.lblMitteStiche_1) + CInt(frmSpielBogen.lblMitteEuL_1) + CInt(frmSpielBogen.lblMitteRK_1) + CInt(frmSpielBogen.lblMitteLorum_1)
            frmSpielBogen.lblRechtsSumme_1 = CInt(frmSpielBogen.lblRechtsRot_1) + CInt(frmSpielBogen.lblRechtsOber_1) + CInt(frmSpielBogen.lblRechtsStiche_1) + CInt(frmSpielBogen.lblRechtsEuL_1) + CInt(frmSpielBogen.lblRechtsRK_1) + CInt(frmSpielBogen.lblRechtsLorum_1)
            frmSpielBogen.lblSpielerSumme_1 = CInt(frmSpielBogen.lblSpielerRot_1) + CInt(frmSpielBogen.lblSpielerOber_1) + CInt(frmSpielBogen.lblSpielerStiche_1) + CInt(frmSpielBogen.lblSpielerEuL_1) + CInt(frmSpielBogen.lblSpielerRK_1) + CInt(frmSpielBogen.lblSpielerLorum_1)
            frmSpielBogen.lblLinksZwS_1 = frmSpielBogen.lblLinksSumme_1
            frmSpielBogen.lblMitteZwS_1 = frmSpielBogen.lblMitteSumme_1
            frmSpielBogen.lblRechtsZwS_1 = frmSpielBogen.lblRechtsSumme_1
            frmSpielBogen.lblSpielerZwS_1 = frmSpielBogen.lblSpielerSumme_1
        ElseIf Rundez�hler = 2 Then
            frmSpielBogen.lblLinksLorum_2 = LinksPunkte
            frmSpielBogen.lblMitteLorum_2 = MittePunkte
            frmSpielBogen.lblRechtsLorum_2 = RechtsPunkte
            frmSpielBogen.lblSpielerLorum_2 = SpielerPunkte
            frmSpielBogen.lblLinksSumme_2 = CInt(frmSpielBogen.lblLinksRot_2) + CInt(frmSpielBogen.lblLinksOber_2) + CInt(frmSpielBogen.lblLinksStiche_2) + CInt(frmSpielBogen.lblLinksEuL_2) + CInt(frmSpielBogen.lblLinksRK_2) + CInt(frmSpielBogen.lblLinksLorum_2)
            frmSpielBogen.lblMitteSumme_2 = CInt(frmSpielBogen.lblMitteRot_2) + CInt(frmSpielBogen.lblMItteOber_2) + CInt(frmSpielBogen.lblMitteStiche_2) + CInt(frmSpielBogen.lblMitteEuL_2) + CInt(frmSpielBogen.lblMitteRK_2) + CInt(frmSpielBogen.lblMitteLorum_2)
            frmSpielBogen.lblRechtsSumme_2 = CInt(frmSpielBogen.lblRechtsRot_2) + CInt(frmSpielBogen.lblRechtsOber_2) + CInt(frmSpielBogen.lblRechtsStiche_2) + CInt(frmSpielBogen.lblRechtsEuL_2) + CInt(frmSpielBogen.lblRechtsRK_2) + CInt(frmSpielBogen.lblRechtsLorum_2)
            frmSpielBogen.lblSpielerSumme_2 = CInt(frmSpielBogen.lblSpielerRot_2) + CInt(frmSpielBogen.lblSpielerOber_2) + CInt(frmSpielBogen.lblSpielerStiche_2) + CInt(frmSpielBogen.lblSpielerEuL_2) + CInt(frmSpielBogen.lblSpielerRK_2) + CInt(frmSpielBogen.lblSpielerLorum_2)
            frmSpielBogen.lblLinksZwS_2 = CInt(frmSpielBogen.lblLinksSumme_2) + CInt(frmSpielBogen.lblLinksZwS_1)
            frmSpielBogen.lblMitteZwS_2 = CInt(frmSpielBogen.lblMitteSumme_2) + CInt(frmSpielBogen.lblMitteZwS_1)
            frmSpielBogen.lblRechtsZwS_2 = CInt(frmSpielBogen.lblRechtsSumme_2) + CInt(frmSpielBogen.lblRechtsZwS_1)
            frmSpielBogen.lblSpielerZwS_2 = CInt(frmSpielBogen.lblSpielerSumme_2) + CInt(frmSpielBogen.lblSpielerZwS_1)
        ElseIf Rundez�hler = 3 Then
            frmSpielBogen.lblLinksLorum_3 = LinksPunkte
            frmSpielBogen.lblMitteLorum_3 = MittePunkte
            frmSpielBogen.lblRechtsLorum_3 = RechtsPunkte
            frmSpielBogen.lblSpielerLorum_3 = SpielerPunkte
            frmSpielBogen.lblLinksSumme_3 = CInt(frmSpielBogen.lblLinksRot_3) + CInt(frmSpielBogen.lblLinksOber_3) + CInt(frmSpielBogen.lblLinksStiche_3) + CInt(frmSpielBogen.lblLinksEuL_3) + CInt(frmSpielBogen.lblLinksRK_3) + CInt(frmSpielBogen.lblLinksLorum_3)
            frmSpielBogen.lblMitteSumme_3 = CInt(frmSpielBogen.lblMitteRot_3) + CInt(frmSpielBogen.lblMItteOber_3) + CInt(frmSpielBogen.lblMitteStiche_3) + CInt(frmSpielBogen.lblMitteEuL_3) + CInt(frmSpielBogen.lblMitteRK_3) + CInt(frmSpielBogen.lblMitteLorum_3)
            frmSpielBogen.lblRechtsSumme_3 = CInt(frmSpielBogen.lblRechtsRot_3) + CInt(frmSpielBogen.lblRechtsOber_3) + CInt(frmSpielBogen.lblRechtsStiche_3) + CInt(frmSpielBogen.lblRechtsEuL_3) + CInt(frmSpielBogen.lblRechtsRK_3) + CInt(frmSpielBogen.lblRechtsLorum_3)
            frmSpielBogen.lblSpielerSumme_3 = CInt(frmSpielBogen.lblSpielerRot_3) + CInt(frmSpielBogen.lblSpielerOber_3) + CInt(frmSpielBogen.lblSpielerStiche_3) + CInt(frmSpielBogen.lblSpielerEuL_3) + CInt(frmSpielBogen.lblSpielerRK_3) + CInt(frmSpielBogen.lblSpielerLorum_3)
            frmSpielBogen.lblLinksZwS_3 = CInt(frmSpielBogen.lblLinksSumme_3) + CInt(frmSpielBogen.lblLinksZwS_2)
            frmSpielBogen.lblMitteZwS_3 = CInt(frmSpielBogen.lblMitteSumme_3) + CInt(frmSpielBogen.lblMitteZwS_2)
            frmSpielBogen.lblRechtsZwS_3 = CInt(frmSpielBogen.lblRechtsSumme_3) + CInt(frmSpielBogen.lblRechtsZwS_2)
            frmSpielBogen.lblSpielerZwS_3 = CInt(frmSpielBogen.lblSpielerSumme_3) + CInt(frmSpielBogen.lblSpielerZwS_2)
        ElseIf Rundez�hler = 4 Then
            frmSpielBogen.lblLinksLorum_4 = LinksPunkte
            frmSpielBogen.lblMitteLorum_4 = MittePunkte
            frmSpielBogen.lblRechtsLorum_4 = RechtsPunkte
            frmSpielBogen.lblSpielerLorum_4 = SpielerPunkte
            frmSpielBogen.lblLinksSumme_4 = CInt(frmSpielBogen.lblLinksRot_4) + CInt(frmSpielBogen.lblLinksOber_4) + CInt(frmSpielBogen.lblLinksStiche_4) + CInt(frmSpielBogen.lblLinksEuL_4) + CInt(frmSpielBogen.lblLinksRK_4) + CInt(frmSpielBogen.lblLinksLorum_4)
            frmSpielBogen.lblMitteSumme_4 = CInt(frmSpielBogen.lblMitteRot_4) + CInt(frmSpielBogen.lblMItteOber_4) + CInt(frmSpielBogen.lblMitteStiche_4) + CInt(frmSpielBogen.lblMitteEuL_4) + CInt(frmSpielBogen.lblMitteRK_4) + CInt(frmSpielBogen.lblMitteLorum_4)
            frmSpielBogen.lblRechtsSumme_4 = CInt(frmSpielBogen.lblRechtsRot_4) + CInt(frmSpielBogen.lblRechtsOber_4) + CInt(frmSpielBogen.lblRechtsStiche_4) + CInt(frmSpielBogen.lblRechtsEuL_4) + CInt(frmSpielBogen.lblRechtsRK_4) + CInt(frmSpielBogen.lblRechtsLorum_4)
            frmSpielBogen.lblSpielerSumme_4 = CInt(frmSpielBogen.lblSpielerRot_4) + CInt(frmSpielBogen.lblSpielerOber_4) + CInt(frmSpielBogen.lblSpielerStiche_4) + CInt(frmSpielBogen.lblSpielerEuL_4) + CInt(frmSpielBogen.lblSpielerRK_4) + CInt(frmSpielBogen.lblSpielerLorum_4)
            frmSpielBogen.lblLinksZwS_4 = CInt(frmSpielBogen.lblLinksSumme_4) + CInt(frmSpielBogen.lblLinksZwS_3)
            frmSpielBogen.lblMitteZwS_4 = CInt(frmSpielBogen.lblMitteSumme_4) + CInt(frmSpielBogen.lblMitteZwS_3)
            frmSpielBogen.lblRechtsZwS_4 = CInt(frmSpielBogen.lblRechtsSumme_4) + CInt(frmSpielBogen.lblRechtsZwS_3)
            frmSpielBogen.lblSpielerZwS_4 = CInt(frmSpielBogen.lblSpielerSumme_4) + CInt(frmSpielBogen.lblSpielerZwS_3)
        End If
    End If
End Sub

Private Sub EintragenBestwerte()
    Dim I As SYSTEMTIME
    Dim strDatum As String
    Dim strMonat As String
    Dim strTag As String
    Dim strStunde As String
    Dim strMinute As String
    Dim strsekunde As String
    Dim strUhrzeit As String
    Dim strPunkte As String
    Dim strSpieler As String
    Dim intNeuerTag As Long

    'Eintragen in RoserlBestwerte.log
    '03.10.2017 Wenn zwei Spieler die gleiche niedrige Punktzahl haben ist strPunkte und strSpieler leer
    'deshalb < ge�ndert in <=
    If frmSpielBogen.lblLinksZwS_4 <= frmSpielBogen.lblMitteZwS_4 And frmSpielBogen.lblLinksZwS_4 <= frmSpielBogen.lblRechtsZwS_4 And _
        frmSpielBogen.lblLinksZwS_4 <= frmSpielBogen.lblSpielerZwS_4 Then
        strPunkte = frmSpielBogen.lblLinksZwS_4
        strSpieler = "Links"
    ElseIf frmSpielBogen.lblMitteZwS_4 <= frmSpielBogen.lblLinksZwS_4 And frmSpielBogen.lblMitteZwS_4 <= frmSpielBogen.lblRechtsZwS_4 And _
        frmSpielBogen.lblMitteZwS_4 <= frmSpielBogen.lblSpielerZwS_4 Then
        strPunkte = frmSpielBogen.lblMitteZwS_4
        strSpieler = "mitte"
    ElseIf frmSpielBogen.lblRechtsZwS_4 <= frmSpielBogen.lblLinksZwS_4 And frmSpielBogen.lblRechtsZwS_4 <= frmSpielBogen.lblMitteZwS_4 And _
        frmSpielBogen.lblRechtsZwS_4 <= frmSpielBogen.lblSpielerZwS_4 Then
        strPunkte = frmSpielBogen.lblRechtsZwS_4
        strSpieler = "Rechts"
    ElseIf frmSpielBogen.lblSpielerZwS_4 <= frmSpielBogen.lblLinksZwS_4 And frmSpielBogen.lblSpielerZwS_4 <= frmSpielBogen.lblMitteZwS_4 And _
        frmSpielBogen.lblSpielerZwS_4 <= frmSpielBogen.lblRechtsZwS_4 Then
        strPunkte = frmSpielBogen.lblSpielerZwS_4
        strSpieler = "Spieler"
    End If
    
    GetSystemTime I
    
    If GMTDiff() = 2 Then                                                               'Gerbing 02.11.2017
        strStunde = I.wHour + 2                                                         'Gerbing 13.09.2017 'Kontrolle bei Ende der Sommerzeit
    Else
        strStunde = I.wHour + 1
    End If
    If strStunde < 24 Then
        intNeuerTag = 0                                                                 'Gerbing 19.09.2017
    Else
        strStunde = strStunde - 24                                                      'Gerbing 19.09.2017
        intNeuerTag = 1                                                                 'Gerbing 19.09.2017
    End If
    If Len(strStunde) = 1 Then
        strStunde = "0" & strStunde
    End If
    strUhrzeit = strStunde & ":"
    strMinute = I.wMinute
    If Len(strMinute) = 1 Then
        strMinute = "0" & strMinute
    End If
    strUhrzeit = strUhrzeit & strMinute & ":"
    strsekunde = I.wSecond
    If Len(strsekunde) = 1 Then
        strsekunde = "0" & strsekunde
    End If
    strUhrzeit = strUhrzeit & strsekunde
    
    strDatum = I.wYear & "-"
    strMonat = I.wMonth
    If Len(strMonat) = 1 Then
        strMonat = "0" & strMonat
    End If
    strDatum = strDatum & strMonat & "-"
    strTag = I.wDay + intNeuerTag
    If Len(strTag) = 1 Then
        strTag = "0" & strTag
    End If
    strDatum = strDatum & strTag

    
    'Neuen Satz in RoserlBestwerte.log eintragen                                        'Gerbing 13.09.2017
    'Set oStream = LogFso.OpenTextFile(filename[, ForAppending[,True[, unicode]])       'True erzeugt die Datei wenn es noch keine gibt
    Set oStream = LogFso.OpenTextFile(App.Path & "\RoserlBestwerte.log", ForAppending, True)
    'Datei Roserl.log Zeile f�r Zeile einlesen und diese Zeile in Grid1 eintragen
    oStream.WriteLine (strDatum & " " & strUhrzeit & vbTab & strSpieler & vbTab & strPunkte)
    oStream.Close
End Sub

' Gibt die Zeitdifferenz zur GMT-Zeitangabe in Stunden zur�ck
' (f�r Deutschland: 1 = Winterzeit, 2 = Sommerzeit)
Private Function GMTDiff() As Long
    Dim nDiff As Long
    Dim st As SYSTEMTIME
 
    ' Systemzeit ermitteln
    GetSystemTime st
    
     ' Zeit-Differenz zur GMT-Zeit in Minuten
     nDiff = DateDiff("n", DateSerial(st.wYear, st.wMonth, st.wDay) + _
        TimeSerial(st.wHour, st.wMinute, st.wSecond), Now)
    
     ' Zeit-Differenz in Stunden
     If (nDiff Mod 60) <> 0 Then nDiff = nDiff - 1
     GMTDiff = nDiff / 60
End Function

Private Sub MoveKarteLorum(Position As Integer)
    'Gerbing 28.07.2020
    'ich musste Facemode = NachOben �ndern in FaceMode = R�ckseite
    'weil ich nicht rausbekomme, warum zu Beginn der Move-Bewegung ein falsches Kartenbild gezeigt wird
    
    Dim I As Integer
    Dim myMaxMove As Integer
    
    On Error GoTo FehlerAusgang
    myMaxMove = iMaxMove
    nDoEvents = 0

    WegX = ZielPunktLeftLorum - StartPunktXLorum
    WegY = ZielPunktTopLorum - StartPunktYLorum
    AbsWegX = Abs(WegX)
    AbsWegY = Abs(WegY)
    If AbsWegX < AbsWegY Then
        XMulti = AbsWegX / AbsWegY
        YMulti = 1
    ElseIf AbsWegY < AbsWegX Then
        YMulti = AbsWegY / AbsWegX
        XMulti = 1
    ElseIf AbsWegY = AbsWegX Then
        XMulti = 1
        YMulti = 1
    End If
    For I = 1 To myMaxMove
        If WegX < 0 And WegY > 0 Then
            If strAbleger = "Links" Then
                cdLinks(Position).FaceMode = R�ckseite
                cdLinks(Position).ZOrder
                cdLinks(Position).Move StartPunktXLorum - I * XMulti, StartPunktYLorum + I * YMulti      'Bewegungsanimation nach unten links
            ElseIf strAbleger = "Mitte" Then
                cdMitte(Position).FaceMode = R�ckseite
                cdMitte(Position).ZOrder
                cdMitte(Position).Move StartPunktXLorum - I * XMulti, StartPunktYLorum + I * YMulti      'Bewegungsanimation nach unten links
            ElseIf strAbleger = "Rechts" Then
                cdRechts(Position).FaceMode = R�ckseite
                cdRechts(Position).ZOrder
                cdRechts(Position).Move StartPunktXLorum - I * XMulti, StartPunktYLorum + I * YMulti     'Bewegungsanimation nach unten links
            ElseIf strAbleger = "Spieler" Then
                cdSpieler(Position).ZOrder
                cdSpieler(Position).Move StartPunktXLorum - I * XMulti, StartPunktYLorum + I * YMulti    'Bewegungsanimation nach unten links
            End If
        ElseIf WegX > 0 And WegY > 0 Then
            If strAbleger = "Links" Then
                cdLinks(Position).FaceMode = R�ckseite
                cdLinks(Position).ZOrder
                cdLinks(Position).Move StartPunktXLorum + I * XMulti, StartPunktYLorum + I * YMulti      'Bewegungsanimation nach unten rechts
            ElseIf strAbleger = "Mitte" Then
                cdMitte(Position).FaceMode = R�ckseite
                cdMitte(Position).ZOrder
                cdMitte(Position).Move StartPunktXLorum + I * XMulti, StartPunktYLorum + I * YMulti      'Bewegungsanimation nach unten rechts
            ElseIf strAbleger = "Rechts" Then
                cdRechts(Position).FaceMode = R�ckseite
                cdRechts(Position).ZOrder
                cdRechts(Position).Move StartPunktXLorum + I * XMulti, StartPunktYLorum + I * YMulti     'Bewegungsanimation nach unten rechts
            ElseIf strAbleger = "Spieler" Then
                cdSpieler(Position).ZOrder
                cdSpieler(Position).Move StartPunktXLorum + I * XMulti, StartPunktYLorum + I * YMulti    'Bewegungsanimation nach unten rechts
            End If
        ElseIf WegX > 0 And WegY < 0 Then
            If strAbleger = "Links" Then
                cdLinks(Position).FaceMode = R�ckseite
                cdLinks(Position).ZOrder
                cdLinks(Position).Move StartPunktXLorum + I * XMulti, StartPunktYLorum - I * YMulti      'Bewegungsanimation nach oben rechts
            ElseIf strAbleger = "Mitte" Then
                cdMitte(Position).FaceMode = R�ckseite
                cdMitte(Position).ZOrder
                cdMitte(Position).Move StartPunktXLorum + I * XMulti, StartPunktYLorum - I * YMulti      'Bewegungsanimation nach oben rechts
            ElseIf strAbleger = "Rechts" Then
                cdRechts(Position).FaceMode = R�ckseite
                cdRechts(Position).ZOrder
                cdRechts(Position).Move StartPunktXLorum + I * XMulti, StartPunktYLorum - I * YMulti     'Bewegungsanimation nach oben rechts
            ElseIf strAbleger = "Spieler" Then
                cdSpieler(Position).ZOrder
                cdSpieler(Position).Move StartPunktXLorum + I * XMulti, StartPunktYLorum - I * YMulti    'Bewegungsanimation nach oben rechts
            End If
        ElseIf WegX < 0 And WegY < 0 Then
            If strAbleger = "Links" Then
                cdLinks(Position).FaceMode = R�ckseite
                cdLinks(Position).ZOrder
                cdLinks(Position).Move StartPunktXLorum - I * XMulti, StartPunktYLorum - I * YMulti      'Bewegungsanimation nach oben links
            ElseIf strAbleger = "Mitte" Then
                cdMitte(Position).FaceMode = R�ckseite
                cdMitte(Position).ZOrder
                cdMitte(Position).Move StartPunktXLorum - I * XMulti, StartPunktYLorum - I * YMulti      'Bewegungsanimation nach oben links
            ElseIf strAbleger = "Rechts" Then
                cdRechts(Position).FaceMode = R�ckseite
                cdRechts(Position).ZOrder
                cdRechts(Position).Move StartPunktXLorum - I * XMulti, StartPunktYLorum - I * YMulti     'Bewegungsanimation nach oben links
            ElseIf strAbleger = "Spieler" Then
                cdSpieler(Position).ZOrder
                cdSpieler(Position).Move StartPunktXLorum - I * XMulti, StartPunktYLorum - I * YMulti    'Bewegungsanimation nach oben links
            End If
        End If
        Sleep (MillisecSleepLorum)
        If nDoEvents <> iMaxDoEvents Then
            nDoEvents = nDoEvents + 1
        Else
            DoEvents
            nDoEvents = 0
        End If
    Next I
FehlerAusgang:
End Sub

Private Function HatKeinerMehr(computercards() As cRoserl, SRGEPosition As Integer)
    'rc=0 die Farbe gibts noch
    'rc=1 die Farbe hat keiner mehr
    Dim z�hlerAK As Integer                                             'z�hler Ausgespielte Karten
    Dim z�hlerMK As Integer                                             'z�hler Meine Karten
    Dim I As Integer
    
    'Kontrolliere ob eine Farbe gespielt werden soll, von der bereits alle Karten ausgespielt sind
    'das ist dann der Fall wenn z�hlerAK plus z�hlerMK = 8
    z�hlerAK = 0
    For I = 1 To AKIndex
        If computercards(SRGEPosition).cType = AusgespielteKarten(I).cType Then
            z�hlerAK = z�hlerAK + 1
        End If
    Next I
    z�hlerMK = 0
    For I = 1 To UBound(computercards)
        If computercards(I).cType = computercards(SRGEPosition).cType Then
            z�hlerMK = z�hlerMK + 1
        End If
    Next I
    If z�hlerAK + z�hlerMK = 8 Then
        HatKeinerMehr = 1
    Else
        HatKeinerMehr = 0
    End If
End Function

Private Function VorbehandlungRotOderStiche(computercards() As cRoserl, PlayerTurn As Players, AKType As Integer)
    'Wenn ich der 4.Zugeber(der letzte eines Stiches) bin, muss ich versuchen
    'geradeso unterhalb der h�chsten Karte zu bleiben, die auf dem Tisch liegt(3 liegen auf dem Tisch), nicht einfach die niedrigste zugeben
    Dim H�chsterWert As Integer
    Dim NiedrigsterWert As Integer
    Dim I As Integer
    Dim MyPosition As Integer
    Dim SaveMyPosition As Integer
    
    'Welches ist die h�chste Karte, die auf dem Tisch liegt?
    If PlayerTurn = Links Then
        H�chsterWert = MitteGespielt.CardValue
        If RechtsGespielt.CardValue > H�chsterWert And AKType = RechtsGespielt.CardType Then
            H�chsterWert = RechtsGespielt.CardValue
        End If
        If SpielerGespielt.CardValue > H�chsterWert And AKType = SpielerGespielt.CardType Then
            H�chsterWert = SpielerGespielt.CardValue
        End If
    ElseIf PlayerTurn = Mitte Then
        H�chsterWert = RechtsGespielt.CardValue
        If SpielerGespielt.CardValue > H�chsterWert And AKType = SpielerGespielt.CardType Then
            H�chsterWert = SpielerGespielt.CardValue
        End If
        If LinksGespielt.CardValue > H�chsterWert And AKType = LinksGespielt.CardType Then
            H�chsterWert = LinksGespielt.CardValue
        End If
    ElseIf PlayerTurn = Rechts Then
        H�chsterWert = SpielerGespielt.CardValue
        If H�chsterWert < LinksGespielt.CardValue And AKType = LinksGespielt.CardType Then
            H�chsterWert = LinksGespielt.CardValue
        End If
        If H�chsterWert < MitteGespielt.CardValue And AKType = MitteGespielt.CardType Then
            H�chsterWert = MitteGespielt.CardValue
        End If
    End If
    'H�chsterWert = h�chste Karte, die auf dem Tisch liegt(3 liegen auf dem Tisch)
    'jetzt geradeso unterhalb der h�chsten Karte bleiben
    NiedrigsterWert = 0
    For I = 1 To UBound(computercards)
        If AKType = computercards(I).cType Then
            MyPosition = I
            If computercards(I).cValue >= NiedrigsterWert And computercards(I).cValue <= H�chsterWert Then
                NiedrigsterWert = computercards(I).cValue
                MyPosition = I
                SaveMyPosition = MyPosition
            Else
                If SaveMyPosition <> 0 Then
                    MyPosition = SaveMyPosition
                End If
            End If
        End If
    Next I
    VorbehandlungRotOderStiche = MyPosition
    Exit Function
End Function

Private Function VorbehandlungOber(computercards() As cRoserl, PlayerTurn As Players, AKType As Integer)
    'Wenn der Ober noch nicht gespielt ist und nicht auf dem Tisch liegt
    'und wenn ich der 4.Zugeber(der letzte eines Stiches) bin, kann ich K�nig oder Ass zugeben
    'aber nicht den Ober                                            'Gerbing 28.11.2017
    'nicht einfach die niedrigste zugeben, sondern die h�chste suchen
    Dim H�chsterWert As Integer
    Dim I As Integer
    Dim MyPosition As Integer

    H�chsterWert = 0
    For I = 1 To UBound(computercards)
        If AKType = computercards(I).cType Then
            If computercards(I).cValue = 8 Or computercards(I).cValue = 7 Then          '8=Ass 7=K�nig Gerbing 28.11.2017
                'MyPosition = I
                If H�chsterWert < computercards(I).cValue Then
                    H�chsterWert = computercards(I).cValue
                    MyPosition = I
                End If
            End If
        End If
    Next I
    VorbehandlungOber = MyPosition
End Function

Private Function VorbehandlungRoterK�nig(computercards() As cRoserl, PlayerTurn As Players, AKType As Integer)
    'Wenn der RoterK�nig noch nicht gespielt ist und nicht auf dem Tisch liegt
    'und wenn ich der 4.Zugeber(der letzte eines Stiches) bin, kann ich Ass zugeben
    'nicht einfach die niedrigste zugeben, sondern die h�chste suchen
    'aber nicht den Roten K�nig                                         'Gerbing 28.11.2017
    Dim H�chsterWert As Integer
    Dim I As Integer
    Dim MyPosition As Integer

    H�chsterWert = 0
    For I = 1 To UBound(computercards)
        If computercards(I).cValue <> 7 Then                            '7=K�nig Gerbing 28.11.2017
            If AKType = computercards(I).cType Then
                'MyPosition = I
                If H�chsterWert < computercards(I).cValue Then
                    H�chsterWert = computercards(I).cValue
                    MyPosition = I
                End If
            End If
        End If
    Next I
    VorbehandlungRoterK�nig = MyPosition
End Function

Private Sub TimerEinsammlerLinks_Timer()
    Aufrufz�hlerEinsammler = Aufrufz�hlerEinsammler + 1
    If Aufrufz�hlerEinsammler <> iMaxMove Then
        RechtsGespielt.Move EndpunktXRechts - Aufrufz�hlerEinsammler * iMulti, EndPunktYRechts               'Bewegungsanimation nach links
        LinksGespielt.Move EndpunktXLinks - Aufrufz�hlerEinsammler * iMulti, EndPunktYLinks                  'Bewegungsanimation nach links
        MitteGespielt.Move EndpunktXMitte - Aufrufz�hlerEinsammler * iMulti, EndPunktYMitte                  'Bewegungsanimation nach links
        SpielerGespielt.Move EndpunktXSpieler - Aufrufz�hlerEinsammler * iMulti, EndPunktYSpieler            'Bewegungsanimation nach links
    Else
        TimerEinsammlerLinks.Enabled = False
    End If
End Sub

Private Sub TimerEinsammlerMitte_Timer()
    Aufrufz�hlerEinsammler = Aufrufz�hlerEinsammler + 1
    If Aufrufz�hlerEinsammler <> iMaxMove Then
        RechtsGespielt.Move EndpunktXRechts, EndPunktYRechts - Aufrufz�hlerEinsammler * iMulti               'Bewegungsanimation nach oben
        LinksGespielt.Move EndpunktXLinks, EndPunktYLinks - Aufrufz�hlerEinsammler * iMulti                  'Bewegungsanimation nach oben
        MitteGespielt.Move EndpunktXMitte, EndPunktYMitte - Aufrufz�hlerEinsammler * iMulti                  'Bewegungsanimation nach oben
        SpielerGespielt.Move EndpunktXSpieler, EndPunktYSpieler - Aufrufz�hlerEinsammler * iMulti            'Bewegungsanimation nach oben
    Else
        TimerEinsammlerMitte.Enabled = False
    End If
End Sub

Private Sub TimerEinsammlerRechts_Timer()
    Aufrufz�hlerEinsammler = Aufrufz�hlerEinsammler + 1
    If Aufrufz�hlerEinsammler <> iMaxMove Then
        RechtsGespielt.Move EndpunktXRechts + Aufrufz�hlerEinsammler * iMulti, EndPunktYRechts               'Bewegungsanimation nach rechts
        LinksGespielt.Move EndpunktXLinks + Aufrufz�hlerEinsammler * iMulti, EndPunktYLinks                  'Bewegungsanimation nach rechts
        MitteGespielt.Move EndpunktXMitte + Aufrufz�hlerEinsammler * iMulti, EndPunktYMitte                  'Bewegungsanimation nach rechts
        SpielerGespielt.Move EndpunktXSpieler + Aufrufz�hlerEinsammler * iMulti, EndPunktYSpieler            'Bewegungsanimation nach rechts
    Else
        TimerEinsammlerRechts.Enabled = False
    End If
End Sub

Private Sub TimerEinsammlerSpieler_Timer()
    Aufrufz�hlerEinsammler = Aufrufz�hlerEinsammler + 1
    If Aufrufz�hlerEinsammler <> iMaxMove Then
        RechtsGespielt.Move EndpunktXRechts, EndPunktYRechts + Aufrufz�hlerEinsammler * iMulti               'Bewegungsanimation nach unten
        LinksGespielt.Move EndpunktXLinks, EndPunktYLinks + Aufrufz�hlerEinsammler * iMulti                  'Bewegungsanimation nach unten
        MitteGespielt.Move EndpunktXMitte, EndPunktYMitte + Aufrufz�hlerEinsammler * iMulti                  'Bewegungsanimation nach unten
        SpielerGespielt.Move EndpunktXSpieler, EndPunktYSpieler + Aufrufz�hlerEinsammler * iMulti            'Bewegungsanimation nach unten
    Else
        TimerEinsammlerSpieler.Enabled = False
    End If
End Sub

Private Sub TimerLinksAnimation_Timer()
    Aufrufz�hlerLinks = Aufrufz�hlerLinks + 1
    If Aufrufz�hlerLinks <> iMaxMove Then
        LinksGespielt.Move StartpunktXLinks + Aufrufz�hlerLinks * iMulti, StartPunktYLinks                  'Bewegungsanimation nach rechts
        LinksGespielt.Visible = True                                                        'Gerbing 22.08.2017
    Else
        TimerLinksAnimation.Enabled = False
        EndpunktXLinks = LinksGespielt.Left                                                     'Gerbing 20.09.2017
        EndPunktYLinks = LinksGespielt.Top
    End If
End Sub

Private Sub TimerMitteAnimation_Timer()
    Aufrufz�hlerMitte = Aufrufz�hlerMitte + 1
    If Aufrufz�hlerMitte <> iMaxMove Then
        MitteGespielt.Move StartpunktXMitte, StartPunktYMitte + Aufrufz�hlerMitte * iMulti                  'Bewegungsanimation nach unten
        MitteGespielt.Visible = True                                                        'Gerbing 22.08.2017
    Else
        TimerMitteAnimation.Enabled = False
        EndpunktXMitte = MitteGespielt.Left                                                     'Gerbing 20.09.2017
        EndPunktYMitte = MitteGespielt.Top
    End If
End Sub

Private Sub TimerRechtsAnimation_Timer()
    Aufrufz�hlerRechts = Aufrufz�hlerRechts + 1
    If Aufrufz�hlerRechts <> iMaxMove Then
        RechtsGespielt.Move StartpunktXRechts - Aufrufz�hlerRechts * iMulti, StartPunktYRechts               'Bewegungsanimation nach links
        RechtsGespielt.Visible = True                                                       'Gerbing 22.08.2017
    Else
        TimerRechtsAnimation.Enabled = False
        EndpunktXRechts = RechtsGespielt.Left                                                   'Gerbing 20.09.2017
        EndPunktYRechts = RechtsGespielt.Top
    End If
End Sub

Private Sub TimerSpielerAnimation_Timer()
    Aufrufz�hlerSpieler = Aufrufz�hlerSpieler + 1
    If Aufrufz�hlerSpieler <> iMaxMove Then
        SpielerGespielt.Move StartpunktXSpieler, StartPunktYSpieler - Aufrufz�hlerSpieler * iMulti            'Bewegungsanimation nach oben
        SpielerGespielt.Visible = True
    Else
        TimerSpielerAnimation.Enabled = False
        EndpunktXSpieler = SpielerGespielt.Left
        EndPunktYSpieler = SpielerGespielt.Top
    End If
End Sub

Private Sub TimerWeiterspielenNachAss_Timer()
    TimerWeiterspielenNachAss.Enabled = False                           '20.10.2017
    If PlayerTurn = Links Then
        Call cdLinks_Click(0)
    ElseIf PlayerTurn = Mitte Then
        Call cdMitte_Click(0)
    ElseIf PlayerTurn = Rechts Then
        Call cdRechts_Click(0)
    End If
End Sub

Private Sub SpeichernSpielstand()
    Dim strZeile1 As String * 50
    Dim Dateinummer As Long
    'Mit dem Men� 'Beenden' wird Roserl beendet und gleichzeitig wird der Spielstand in die Bin�rdatei roserl.bin gespeichert.
    'Das sind keine Zwischenspielst�nde, sondern der Spielstand nach jeweils einer Runde.
    'Wenn w�hrend der ersten Runde abgebrochen wird(da gibt es noch kein Runden-Ergebnis), beginnt also Roserl neu
    'Spielstand in die Bin�rdatei roserl.bin speichern
    'Nicht bei Runde1, erst ab Runde2
    If Rundez�hler > 1 Then
        strZeile1 = "Spielstand vom " & Now & ";"
        Dateinummer = FreeFile
        Open App.Path & "\roserl.bin" For Binary Access Write As #Dateinummer
        Put #Dateinummer, , strZeile1
        '---------------------------------------------------------Gerbing 13.06.2018
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksRot_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteRot_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsRot_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerRot_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksOber_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMItteOber_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsOber_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerOber_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksStiche_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteStiche_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsStiche_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerStiche_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksEuL_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteEuL_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsEuL_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerEuL_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksRK_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteRK_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsRK_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerRK_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksLorum_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteLorum_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsLorum_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerLorum_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksSumme_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteSumme_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsSumme_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerSumme_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksZwS_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteZwS_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsZwS_1)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerZwS_1)
    End If
        
    If Rundez�hler > 2 Then
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksRot_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteRot_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsRot_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerRot_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksOber_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMItteOber_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsOber_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerOber_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksStiche_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteStiche_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsStiche_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerStiche_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksEuL_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteEuL_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsEuL_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerEuL_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksRK_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteRK_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsRK_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerRK_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksLorum_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteLorum_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsLorum_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerLorum_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksSumme_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteSumme_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsSumme_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerSumme_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksZwS_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteZwS_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsZwS_2)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerZwS_2)
    End If
        
    If Rundez�hler > 3 Then
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksRot_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteRot_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsRot_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerRot_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksOber_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMItteOber_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsOber_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerOber_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksStiche_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteStiche_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsStiche_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerStiche_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksEuL_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteEuL_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsEuL_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerEuL_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksRK_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteRK_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsRK_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerRK_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksLorum_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteLorum_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsLorum_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerLorum_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksSumme_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteSumme_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsSumme_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerSumme_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblLinksZwS_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblMitteZwS_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblRechtsZwS_3)
        Put #Dateinummer, , CInt(frmSpielBogen.lblSpielerZwS_3)
    End If
    Close #Dateinummer
End Sub
