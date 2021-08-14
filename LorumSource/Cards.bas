Attribute VB_Name = "mdCrazyEights"
Option Explicit

    Enum Players
        Spieler                         '=0
        Links
        Mitte
        Rechts
    End Enum
    
    Enum Spielarten
        Rot                             '=0
        Ober
        Stiche
        ErsterUndLetzterstich
        RoterKönig
        Lorum
    End Enum
    
    Public Initialized As Boolean
    Public PlayerTurn As Players
    Public Spielart As Spielarten
    Public Ausspieler As Players
    Public AusspielerBeiSpielbeginn As Players
    
    Type POINTAPI
        x As Integer
        y As Integer
    End Type
    
    Public Type cRoserl
        cType As CardTypes
        cValue As CardValues
        cGespieltVon As GespieltVonValues
        cEingesammeltVon As EingesammeltVonValues
    End Type
    
    Global Const CARD_WIDTH = 141
    Global Const CARD_HEIGHT = 250
    Global Const CARDS_IN_DECK = 32
    
    Public CardDeck(1 To 32) As cRoserl
    Public AusgespielteKarten(1 To 32) As cRoserl
    
    Public AKIndex As Integer                       'AusgespielteKartenIndex
    Public Rundezähler As Integer
    Public blnSpielWiederholen As Boolean
    Public blnEingesammelt As Boolean                  'Gerbing 13.08.2017
    
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public blnMaloiGetippt As Boolean
    Public CARD_SPACE As Integer
    Public NormalWidth As Integer
    Public LorumWidth As Integer
    Public NormalHeight As Integer
    Public LorumHeight As Integer
    Public ZählerOber As Integer                       'Gerbing 12.09.2017
    Public ZählerRote As Integer                       'Gerbing 12.09.2017
    
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Const HORZRES = 8
    Public Const VERTRES = 10
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    Public StartMillisek As Long
    Public EndMillisek As Long
    'Public blnUnloadFrmSpielBogen As Boolean
    Public blnNeuesSpielBeginnt As Boolean
    Public Mons As clsScreens
    Public MillisecSleep As Integer
    Public MillisecSleepLorum As Integer
    
    Public cdRoserlSchellATop As Long                                           'Gerbing 28.09.2017
    Public cdRoserlSchellALeft As Long
    Public cdRoserlSchellBTop As Long
    Public cdRoserlSchellBLeft As Long
    Public cdRoserlRotATop As Long
    Public cdRoserlRotALeft As Long
    Public cdRoserlRotBTop As Long
    Public cdRoserlRotBLeft As Long
    Public cdRoserlGrünATop As Long
    Public cdRoserlGrünALeft As Long
    Public cdRoserlGrünBTop As Long
    Public cdRoserlGrünBLeft As Long
    Public cdRoserlEichelATop As Long
    Public cdRoserlEichelALeft As Long
    Public cdRoserlEichelBTop As Long
    Public cdRoserlEichelBLeft As Long
    Public cdRoserlSchellAssTop As Long
    Public cdRoserlSchellAssLeft As Long
    Public cdRoserlRotAssTop As Long
    Public cdRoserlRotAssLeft As Long
    Public cdRoserlGrünAssTop As Long
    Public cdRoserlGrünAssLeft As Long
    Public cdRoserlEichelAssTop As Long
    Public cdRoserlEichelAssLeft As Long
    
    Public LinksCards() As cRoserl
    Public MitteCards() As cRoserl
    Public RechtsCards() As cRoserl
    Public SpielerCards() As cRoserl
    Public gblnNichtMischen As Boolean                                          'Gerbing 10.12.2017

Public Sub ResetDeck()
    Dim I As Integer
    
    For I = 1 To 32
        CardDeck(I).cType = (I - 1) \ 8 + 1
        CardDeck(I).cValue = (I - 1) Mod 8 + 1
    Next I
End Sub

Public Function Card(CardValue As CardValues, CardType As CardTypes) As cRoserl
    Card.cType = CardType
    Card.cValue = CardValue
End Function

Public Sub Mischen()
    Dim cT As CardTypes, cV As CardValues, rNum As Integer
    
    'Randomize Timer
    'Randomize                                                       'Gerbing 19.12.2017 auskommentiert 20.03.2020
    Erase CardDeck
    Erase AusgespielteKarten
    For cT = Schell To Eichel
        For cV = Sieben To Ass
            Do
                rNum = Int(Rnd * 32) + 1
            Loop Until CardDeck(rNum).cType = 0
            CardDeck(rNum).cType = cT
            CardDeck(rNum).cValue = cV
        Next cV
    Next cT
End Sub

Public Sub NichtMischen()
'    Const strLinks As String = "11 15 21 25 31 35 41 45 "
'    Const strMitte As String = "12 16 22 26 32 36 42 46 "
'    Const strRechts As String = "13 17 23 27 33 37 43 47 "
'    Const strSpieler As String = "14 18 24 28 34 38 44 48 "

    Const strLinks As String = "12 15 21 25 32 35 18 26 "
    Const strMitte As String = "11 16 28 45 31 36 41 14 "
    Const strRechts As String = "13 17 23 27 33 37 43 47 "
    Const strSpieler As String = "46 42 24 22 34 38 44 48 "


    Dim Gezinkt As String
    Dim strType As String
    Dim strValue As String
    Dim I As Integer
    Dim n As Integer
    
    
    On Error GoTo 0
    Gezinkt = strLinks & strMitte & strRechts & strSpieler
    Erase CardDeck
    Erase AusgespielteKarten
    n = 1
    For I = 1 To 32
        strType = Mid(Gezinkt, n, 1)
        strValue = Mid(Gezinkt, n + 1, 1)
        CardDeck(I).cType = CInt(strType)
        CardDeck(I).cValue = CInt(strValue)
        n = n + 3
    Next I
End Sub

Sub Main()
    Dim strAusspieler As String
    Dim msg As String
    
    Randomize                                                           'Gerbing 20.03.2020
    Set Mons = New clsScreens                                           'Gerbing 05.10.2017
    Call SetStandardWerteLorum
'    FarbeAusspieler = &H8080FF                                          'Hellrot 'Gerbing 15.08.2017
'    FarbeAusspieler = &H0&                                              'schwarz
    FarbeAusspieler = &HFFFF&                                            'gelb
'    FarbeAusspieler = &HFF&                                             'rot
'    Call StandardWerteLorum
    Call StandardWerteNormal
    'blnMaloiGetippt = True                                              'testweise
    Call SpielBogenNull
    'frmSpielBogen.Hide
    frmRoserl.Show
    
    Call ResetDeck
    'gblnNichtMischen = True                                             'Gerbing 10.12.2017
    Call frmRoserl.Kartengeben
    Rundezähler = 1                                                     'Gerbing 17.03.2018
    Call RotBeginnt
'    Call OberBeginnt
'    Call SticheBeginnt
'    Call ErsterUndLetzterBeginnt
'    Call RoterKönigBeginnt
'    Call LorumBeginnt
    Call EinlesenSpielstand                                             'Gerbing 17.03.2018
    blnEingesammelt = True                                              'Gerbing 13.08.2017
    AKIndex = 0
    ZählerOber = 0                                                      'Gerbing 12.09.2017
    ZählerRote = 0                                                      'Gerbing 12.09.2017
    blnNeuesSpielBeginnt = True                                             'Gerbing 06.09.2017
    gstrMsg = "Du musst zum Weiterspielen auf den roten Buchstabe klicken"  'Gerbing 06.09.2017
    frmMsgBox.Show 1                                                        'Gerbing 06.09.2017
End Sub

Private Sub RotBeginnt()
    Call StandardWerteNormal
    Spielart = Rot
    frmRoserl.lblRoserl(Rot).BackColor = &HFF&                          'Buchstabe R wird rot
    AusspielerBeiSpielbeginn = Links
    Ausspieler = Links
End Sub

Private Sub OberBeginnt()
    Call StandardWerteNormal
    Spielart = Ober
    frmRoserl.lblRoserl(Ober).BackColor = &HFF&                          'Buchstabe O wird rot
    AusspielerBeiSpielbeginn = Links
    Ausspieler = Links
End Sub

Private Sub SticheBeginnt()
    Call StandardWerteNormal
    Spielart = Stiche
    frmRoserl.lblRoserl(Stiche).BackColor = &HFF&                          'Buchstabe S wird rot
    AusspielerBeiSpielbeginn = Links
    Ausspieler = Links
End Sub

Private Sub ErsterUndLetzterBeginnt()
    Call StandardWerteNormal
    Spielart = ErsterUndLetzterstich
    frmRoserl.lblRoserl(ErsterUndLetzterstich).BackColor = &HFF&            'Buchstabe E wird rot
    AusspielerBeiSpielbeginn = Links
    Ausspieler = Links
End Sub

Private Sub RoterKönigBeginnt()
    Call StandardWerteNormal
    Spielart = RoterKönig
    frmRoserl.lblRoserl(RoterKönig).BackColor = &HFF&                       'Buchstabe R wird rot
    AusspielerBeiSpielbeginn = Links
    Ausspieler = Links
End Sub

Private Sub LorumBeginnt()
    Call StandardWerteLorum
    Call frmRoserl.Kartengeben                                              'Gerbing 20.09.2017
    Spielart = Lorum
    frmRoserl.lblRoserl(Lorum).BackColor = &HFF&                            'Buchstabe L wird rot
    AusspielerBeiSpielbeginn = Links
    Ausspieler = Links
    frmRoserl.PlayerTurnRoserl = Links
End Sub

Public Function AnpassenNutzerWunsch(Form)                                              'Gerbing 11.03.2017
    Dim I As Long
    
    'PublicCheckForDPI = 1 = klein
    'PublicCheckForDPI = 2 = mittel
    'PublicCheckForDPI = 3 = groß
    On Error Resume Next
    If PublicCheckForDPI = 1 Then
        For I = 0 To Form.Controls.Count - 1
            Form.Controls(I).Font.Bold = False
            Form.Controls(I).FontName = frmRoserl.txtFont.FontName
            Form.Controls(I).Font.Size = 8
            ''Debug.Print Form.Controls(i).Name
        Next I
    End If
    If PublicCheckForDPI = 2 Then
        For I = 0 To Form.Controls.Count - 1
            Form.Controls(I).Font.Bold = False
            Form.Controls(I).FontName = frmRoserl.txtFont.FontName
            Form.Controls(I).Font.Size = 10
            ''Debug.Print Form.Controls(i).Name
        Next I
    End If
    If PublicCheckForDPI = 3 Then
        For I = 0 To Form.Controls.Count - 1
            Form.Controls(I).Font.Bold = False
            Form.Controls(I).FontName = frmRoserl.txtFont.FontName
            Form.Controls(I).Font.Size = 12
        Next I
    End If
    On Error GoTo 0
End Function

Public Function AnpassenHeadFont(DbGrid)                                                'Gerbing 11.03.2017
    Dim MyAnzeigeGroesse As Integer
    
    MyAnzeigeGroesse = 2
    On Error Resume Next
    DbGrid.HeadFont.Bold = False
    DbGrid.HeadFont.Name = frmRoserl.txtFont.FontName
    If MyAnzeigeGroesse = 1 Then
            DbGrid.HeadFont.Size = 8
            DbGrid.RowHeight = 220
    End If
    If MyAnzeigeGroesse = 2 Then
            DbGrid.HeadFont.Size = 10
            DbGrid.RowHeight = 260
    End If
    If MyAnzeigeGroesse = 3 Then
            DbGrid.HeadFont.Size = 12
            DbGrid.RowHeight = 300
    End If
    On Error GoTo 0
End Function

Public Sub StandardWerteNormal()
    CARD_SPACE = 150
    'CARD_SPACE = 135
    NormalHeight = 251
    NormalWidth = 141
End Sub

Public Sub StandardWerteLorum()
    Dim I As Integer
'    CARD_SPACE = 75
'    NormalHeight = 125
'    NormalWidth = 70
    CARD_SPACE = 95
    NormalHeight = 159
    NormalWidth = 90
    'alle Lorum-Karten-Positionen sind in SetStandardWerteLorum gespeichert worden und müssen jetzt wiederhergestellt werden
    'Bei jedem neuen Lorum-Spiel ist zwar die Anzahl aller Lorum-Ablegekarten auf eine zurückgesetzt worden in AlleKartenEntladen,
    'aber deren .Left-Position ist verschoben worden
    frmRoserl.cdRoserlSchellA(I).Left = cdRoserlSchellALeft                 'Gerbing 28.09.2017
    frmRoserl.cdRoserlSchellB(I).Left = cdRoserlSchellBLeft
    frmRoserl.cdRoserlRotA(I).Left = cdRoserlRotALeft
    frmRoserl.cdRoserlRotB(I).Left = cdRoserlRotBLeft
    frmRoserl.cdRoserlGrünA(I).Left = cdRoserlGrünALeft
    frmRoserl.cdRoserlGrünB(I).Left = cdRoserlGrünBLeft
    frmRoserl.cdRoserlEichelA(I).Left = cdRoserlEichelALeft
    frmRoserl.cdRoserlEichelB(I).Left = cdRoserlEichelBLeft
    frmRoserl.cdRoserlSchellA(I).Left = cdRoserlSchellALeft
    frmRoserl.cdRoserlSchellA(I).Left = cdRoserlSchellALeft
    frmRoserl.cdRoserlSchellA(I).Left = cdRoserlSchellALeft
    frmRoserl.cdRoserlSchellA(I).Left = cdRoserlSchellALeft
End Sub

Private Sub SetStandardWerteLorum()                                         'Gerbing 28.09.2017
    cdRoserlSchellATop = frmRoserl.cdRoserlSchellA(0).Top
    cdRoserlSchellALeft = frmRoserl.cdRoserlSchellA(0).Left
    cdRoserlSchellBTop = frmRoserl.cdRoserlSchellB(0).Top
    cdRoserlSchellBLeft = frmRoserl.cdRoserlSchellB(0).Left
    cdRoserlRotATop = frmRoserl.cdRoserlRotA(0).Top
    cdRoserlRotALeft = frmRoserl.cdRoserlRotA(0).Left
    cdRoserlRotBTop = frmRoserl.cdRoserlRotB(0).Top
    cdRoserlRotBLeft = frmRoserl.cdRoserlRotB(0).Left
    cdRoserlGrünATop = frmRoserl.cdRoserlGrünA(0).Top
    cdRoserlGrünALeft = frmRoserl.cdRoserlGrünA(0).Left
    cdRoserlGrünBTop = frmRoserl.cdRoserlGrünB(0).Top
    cdRoserlGrünBLeft = frmRoserl.cdRoserlGrünB(0).Left
    cdRoserlEichelATop = frmRoserl.cdRoserlEichelA(0).Top
    cdRoserlEichelALeft = frmRoserl.cdRoserlEichelA(0).Left
    cdRoserlEichelBTop = frmRoserl.cdRoserlEichelB(0).Top
    cdRoserlEichelBLeft = frmRoserl.cdRoserlEichelB(0).Left
    cdRoserlSchellAssTop = frmRoserl.cdRoserlSchellAss(0).Top
    cdRoserlSchellAssLeft = frmRoserl.cdRoserlSchellAss(0).Left
    cdRoserlRotAssTop = frmRoserl.cdRoserlRotAss(0).Top
    cdRoserlRotAssLeft = frmRoserl.cdRoserlRotAss(0).Left
    cdRoserlGrünAssTop = frmRoserl.cdRoserlGrünAss(0).Top
    cdRoserlGrünAssLeft = frmRoserl.cdRoserlGrünAss(0).Left
    cdRoserlEichelAssTop = frmRoserl.cdRoserlEichelAss(0).Top
    cdRoserlEichelAssLeft = frmRoserl.cdRoserlEichelAss(0).Left
End Sub

Public Sub SpielBogenNull()
    frmSpielBogen.lblLinksRot_1 = 0
    frmSpielBogen.lblMitteRot_1 = 0
    frmSpielBogen.lblRechtsRot_1 = 0
    frmSpielBogen.lblSpielerRot_1 = 0
    frmSpielBogen.lblLinksRot_2 = 0
    frmSpielBogen.lblMitteRot_2 = 0
    frmSpielBogen.lblRechtsRot_2 = 0
    frmSpielBogen.lblSpielerRot_2 = 0
    frmSpielBogen.lblLinksRot_3 = 0
    frmSpielBogen.lblMitteRot_3 = 0
    frmSpielBogen.lblRechtsRot_3 = 0
    frmSpielBogen.lblSpielerRot_3 = 0
    frmSpielBogen.lblLinksRot_4 = 0
    frmSpielBogen.lblMitteRot_4 = 0
    frmSpielBogen.lblRechtsRot_4 = 0
    frmSpielBogen.lblSpielerRot_4 = 0
    '----------------------------------
    frmSpielBogen.lblLinksOber_1 = 0
    frmSpielBogen.lblMItteOber_1 = 0
    frmSpielBogen.lblRechtsOber_1 = 0
    frmSpielBogen.lblSpielerOber_1 = 0
    frmSpielBogen.lblLinksOber_2 = 0
    frmSpielBogen.lblMItteOber_2 = 0
    frmSpielBogen.lblRechtsOber_2 = 0
    frmSpielBogen.lblSpielerOber_2 = 0
    frmSpielBogen.lblLinksOber_3 = 0
    frmSpielBogen.lblMItteOber_3 = 0
    frmSpielBogen.lblRechtsOber_3 = 0
    frmSpielBogen.lblSpielerOber_3 = 0
    frmSpielBogen.lblLinksOber_4 = 0
    frmSpielBogen.lblMItteOber_4 = 0
    frmSpielBogen.lblRechtsOber_4 = 0
    frmSpielBogen.lblSpielerOber_4 = 0
    '----------------------------------
    frmSpielBogen.lblLinksStiche_1 = 0
    frmSpielBogen.lblMitteStiche_1 = 0
    frmSpielBogen.lblRechtsStiche_1 = 0
    frmSpielBogen.lblSpielerStiche_1 = 0
    frmSpielBogen.lblLinksStiche_2 = 0
    frmSpielBogen.lblMitteStiche_2 = 0
    frmSpielBogen.lblRechtsStiche_2 = 0
    frmSpielBogen.lblSpielerStiche_2 = 0
    frmSpielBogen.lblLinksStiche_3 = 0
    frmSpielBogen.lblMitteStiche_3 = 0
    frmSpielBogen.lblRechtsStiche_3 = 0
    frmSpielBogen.lblSpielerStiche_3 = 0
    frmSpielBogen.lblLinksStiche_4 = 0
    frmSpielBogen.lblMitteStiche_4 = 0
    frmSpielBogen.lblRechtsStiche_4 = 0
    frmSpielBogen.lblSpielerStiche_4 = 0
    '----------------------------------
    frmSpielBogen.lblLinksEuL_1 = 0
    frmSpielBogen.lblMitteEuL_1 = 0
    frmSpielBogen.lblRechtsEuL_1 = 0
    frmSpielBogen.lblSpielerEuL_1 = 0
    frmSpielBogen.lblLinksEuL_2 = 0
    frmSpielBogen.lblMitteEuL_2 = 0
    frmSpielBogen.lblRechtsEuL_2 = 0
    frmSpielBogen.lblSpielerEuL_2 = 0
    frmSpielBogen.lblLinksEuL_3 = 0
    frmSpielBogen.lblMitteEuL_3 = 0
    frmSpielBogen.lblRechtsEuL_3 = 0
    frmSpielBogen.lblSpielerEuL_3 = 0
    frmSpielBogen.lblLinksEuL_4 = 0
    frmSpielBogen.lblMitteEuL_4 = 0
    frmSpielBogen.lblRechtsEuL_4 = 0
    frmSpielBogen.lblSpielerEuL_4 = 0
    '----------------------------------
    frmSpielBogen.lblLinksRK_1 = 0
    frmSpielBogen.lblMitteRK_1 = 0
    frmSpielBogen.lblRechtsRK_1 = 0
    frmSpielBogen.lblSpielerRK_1 = 0
    frmSpielBogen.lblLinksRK_2 = 0
    frmSpielBogen.lblMitteRK_2 = 0
    frmSpielBogen.lblRechtsRK_2 = 0
    frmSpielBogen.lblSpielerRK_2 = 0
    frmSpielBogen.lblLinksRK_3 = 0
    frmSpielBogen.lblMitteRK_3 = 0
    frmSpielBogen.lblRechtsRK_3 = 0
    frmSpielBogen.lblSpielerRK_3 = 0
    frmSpielBogen.lblLinksRK_4 = 0
    frmSpielBogen.lblMitteRK_4 = 0
    frmSpielBogen.lblRechtsRK_4 = 0
    frmSpielBogen.lblSpielerRK_4 = 0
    '----------------------------------
    frmSpielBogen.lblLinksLorum_1 = 0
    frmSpielBogen.lblMitteLorum_1 = 0
    frmSpielBogen.lblRechtsLorum_1 = 0
    frmSpielBogen.lblSpielerLorum_1 = 0
    frmSpielBogen.lblLinksLorum_2 = 0
    frmSpielBogen.lblMitteLorum_2 = 0
    frmSpielBogen.lblRechtsLorum_2 = 0
    frmSpielBogen.lblSpielerLorum_2 = 0
    frmSpielBogen.lblLinksLorum_3 = 0
    frmSpielBogen.lblMitteLorum_3 = 0
    frmSpielBogen.lblRechtsLorum_3 = 0
    frmSpielBogen.lblSpielerLorum_3 = 0
    frmSpielBogen.lblLinksLorum_4 = 0
    frmSpielBogen.lblMitteLorum_4 = 0
    frmSpielBogen.lblRechtsLorum_4 = 0
    frmSpielBogen.lblSpielerLorum_4 = 0
    '----------------------------------
    frmSpielBogen.lblLinksSumme_1 = 0
    frmSpielBogen.lblLinksZwS_1 = 0
    frmSpielBogen.lblMitteSumme_1 = 0
    frmSpielBogen.lblMitteZwS_1 = 0
    frmSpielBogen.lblRechtsSumme_1 = 0
    frmSpielBogen.lblRechtsZwS_1 = 0
    frmSpielBogen.lblSpielerSumme_1 = 0
    frmSpielBogen.lblSpielerZwS_1 = 0
    '--------------------------------
    frmSpielBogen.lblLinksSumme_2 = 0
    frmSpielBogen.lblLinksZwS_2 = 0
    frmSpielBogen.lblMitteSumme_2 = 0
    frmSpielBogen.lblMitteZwS_2 = 0
    frmSpielBogen.lblRechtsSumme_2 = 0
    frmSpielBogen.lblRechtsZwS_2 = 0
    frmSpielBogen.lblSpielerSumme_2 = 0
    frmSpielBogen.lblSpielerZwS_2 = 0
    '--------------------------------
    frmSpielBogen.lblLinksSumme_3 = 0
    frmSpielBogen.lblLinksZwS_3 = 0
    frmSpielBogen.lblMitteSumme_3 = 0
    frmSpielBogen.lblMitteZwS_3 = 0
    frmSpielBogen.lblRechtsSumme_3 = 0
    frmSpielBogen.lblRechtsZwS_3 = 0
    frmSpielBogen.lblSpielerSumme_3 = 0
    frmSpielBogen.lblSpielerZwS_3 = 0
    '--------------------------------
    frmSpielBogen.lblLinksSumme_4 = 0
    frmSpielBogen.lblLinksZwS_4 = 0
    frmSpielBogen.lblMitteSumme_4 = 0
    frmSpielBogen.lblMitteZwS_4 = 0
    frmSpielBogen.lblRechtsSumme_4 = 0
    frmSpielBogen.lblRechtsZwS_4 = 0
    frmSpielBogen.lblSpielerSumme_4 = 0
    frmSpielBogen.lblSpielerZwS_4 = 0
End Sub

Public Sub AusspielerFett()                                        'Gerbing 05.09.2017
    Dim strAusspieler As String

    If Ausspieler = Links Then
        strAusspieler = "Links"
        frmRoserl.lblLinks.Font.Bold = True                           'Gerbing 15.08.2017
        frmRoserl.lblLinks.ForeColor = FarbeAusspieler                           'schwarz
        frmRoserl.lblSpieler.ForeColor = &HFFFFFF                     'weiß
        frmRoserl.lblMitte.ForeColor = &HFFFFFF
        frmRoserl.lblRechts.ForeColor = &HFFFFFF
        frmRoserl.lblMitte.Font.Bold = False
        frmRoserl.lblRechts.Font.Bold = False
        frmRoserl.lblSpieler.Font.Bold = False
    End If
    If Ausspieler = Mitte Then
        strAusspieler = "Mitte"
        frmRoserl.lblMitte.Font.Bold = True                           'Gerbing 15.08.2017
        frmRoserl.lblMitte.ForeColor = FarbeAusspieler                           'schwarz
        frmRoserl.lblLinks.ForeColor = &HFFFFFF                       'weiß
        frmRoserl.lblSpieler.ForeColor = &HFFFFFF
        frmRoserl.lblRechts.ForeColor = &HFFFFFF
        frmRoserl.lblLinks.Font.Bold = False
        frmRoserl.lblRechts.Font.Bold = False
        frmRoserl.lblSpieler.Font.Bold = False
    End If
    If Ausspieler = Rechts Then
        strAusspieler = "Rechts"
        frmRoserl.lblRechts.Font.Bold = True                          'Gerbing 15.08.2017
        frmRoserl.lblRechts.ForeColor = FarbeAusspieler                          'schwarz
        frmRoserl.lblLinks.ForeColor = &HFFFFFF                       'weiß
        frmRoserl.lblMitte.ForeColor = &HFFFFFF
        frmRoserl.lblSpieler.ForeColor = &HFFFFFF
        frmRoserl.lblLinks.Font.Bold = False
        frmRoserl.lblMitte.Font.Bold = False
        frmRoserl.lblSpieler.Font.Bold = False
    End If
    If Ausspieler = Spieler Then
        strAusspieler = "Spieler"
        frmRoserl.lblSpieler.Font.Bold = True                         'Gerbing 15.08.2017
        frmRoserl.lblSpieler.ForeColor = FarbeAusspieler                         'schwarz
        frmRoserl.lblLinks.ForeColor = &HFFFFFF                       'weiß
        frmRoserl.lblMitte.ForeColor = &HFFFFFF
        frmRoserl.lblRechts.ForeColor = &HFFFFFF
        frmRoserl.lblLinks.Font.Bold = False
        frmRoserl.lblMitte.Font.Bold = False
        frmRoserl.lblRechts.Font.Bold = False
    End If
    If Spielart = Lorum Then
        gstrMsg = "Wir spielen Lorum. Ausspieler ist " & strAusspieler & "." & vbNewLine
        gstrMsg = gstrMsg & "Klicke auf den Kartenstapel des Ausspielers"
        frmMsgBox.Show 1
        Exit Sub
    End If
    If Ausspieler = Links Then
        Call frmRoserl.LinksPlay                                            'Beim ersten Spiel beginnt Links auszuspielen
        Call frmRoserl.MittePlay
        Call frmRoserl.RechtsPlay
    ElseIf Ausspieler = Mitte Then
        Call frmRoserl.MittePlay
        Call frmRoserl.RechtsPlay
    ElseIf Ausspieler = Rechts Then
        Call frmRoserl.RechtsPlay
    End If

End Sub

Private Sub EinlesenSpielstand()
'    Gerbing 17.03.2018
'    Mit roserl.bin ist es möglich, bei einem abgebrochenen Spielstand weiterzuspielen. Wenn das Programm eine roserl.bin entdeckt,
'    wird der Nutzer gefragt, ob er mit dem Spielstand zu Beginn der abgebrochenen Runde weiterspielen will(mit Rot),
'    oder ob er neu beginnen will. Nach dem Einlesen von roserl.bin wird diese gelöscht.
'    Wer in der 2./3./4. Runde der Rot-Ausspier ist, steht fest.

    Dim Filename As String
    Dim strZeile1 As String * 50
    '----------------------------------Gerbing 13.06.2018
    Dim LinksRot_1 As Integer
    Dim MitteRot_1 As Integer
    Dim RechtsRot_1 As Integer
    Dim SpielerRot_1 As Integer
    Dim LinksOber_1 As Integer
    Dim MitteOber_1 As Integer
    Dim RechtsOber_1 As Integer
    Dim SpielerOber_1 As Integer
    Dim LinksStiche_1 As Integer
    Dim MitteStiche_1 As Integer
    Dim RechtsStiche_1 As Integer
    Dim SpielerStiche_1 As Integer
    Dim LinksEuL_1 As Integer
    Dim MitteEuL_1 As Integer
    Dim RechtsEuL_1 As Integer
    Dim SpielerEuL_1 As Integer
    Dim LinksRK_1 As Integer
    Dim MitteRK_1 As Integer
    Dim RechtsRK_1 As Integer
    Dim SpielerRK_1 As Integer
    Dim LinksLorum_1 As Integer
    Dim MitteLorum_1 As Integer
    Dim RechtsLorum_1 As Integer
    Dim SpielerLorum_1 As Integer
    Dim LinksSumme_1 As Integer
    Dim MitteSumme_1 As Integer
    Dim RechtsSumme_1 As Integer
    Dim SpielerSumme_1 As Integer
    Dim Links1 As Integer
    Dim Mitte1 As Integer
    Dim Rechts1 As Integer
    Dim Spieler1 As Integer
    
    Dim LinksRot_2 As Integer
    Dim MitteRot_2 As Integer
    Dim RechtsRot_2 As Integer
    Dim SpielerRot_2 As Integer
    Dim LinksOber_2 As Integer
    Dim MitteOber_2 As Integer
    Dim RechtsOber_2 As Integer
    Dim SpielerOber_2 As Integer
    Dim LinksStiche_2 As Integer
    Dim MitteStiche_2 As Integer
    Dim RechtsStiche_2 As Integer
    Dim SpielerStiche_2 As Integer
    Dim LinksEuL_2 As Integer
    Dim MitteEuL_2 As Integer
    Dim RechtsEuL_2 As Integer
    Dim SpielerEuL_2 As Integer
    Dim LinksRK_2 As Integer
    Dim MitteRK_2 As Integer
    Dim RechtsRK_2 As Integer
    Dim SpielerRK_2 As Integer
    Dim LinksLorum_2 As Integer
    Dim MitteLorum_2 As Integer
    Dim RechtsLorum_2 As Integer
    Dim SpielerLorum_2 As Integer
    Dim LinksSumme_2 As Integer
    Dim MitteSumme_2 As Integer
    Dim RechtsSumme_2 As Integer
    Dim SpielerSumme_2 As Integer
    Dim Links2 As Integer
    Dim Mitte2 As Integer
    Dim Rechts2 As Integer
    Dim Spieler2 As Integer
    
    Dim LinksRot_3 As Integer
    Dim MitteRot_3 As Integer
    Dim RechtsRot_3 As Integer
    Dim SpielerRot_3 As Integer
    Dim LinksOber_3 As Integer
    Dim MitteOber_3 As Integer
    Dim RechtsOber_3 As Integer
    Dim SpielerOber_3 As Integer
    Dim LinksStiche_3 As Integer
    Dim MitteStiche_3 As Integer
    Dim RechtsStiche_3 As Integer
    Dim SpielerStiche_3 As Integer
    Dim LinksEuL_3 As Integer
    Dim MitteEuL_3 As Integer
    Dim RechtsEuL_3 As Integer
    Dim SpielerEuL_3 As Integer
    Dim LinksRK_3 As Integer
    Dim MitteRK_3 As Integer
    Dim RechtsRK_3 As Integer
    Dim SpielerRK_3 As Integer
    Dim LinksLorum_3 As Integer
    Dim MitteLorum_3 As Integer
    Dim RechtsLorum_3 As Integer
    Dim SpielerLorum_3 As Integer
    Dim LinksSumme_3 As Integer
    Dim MitteSumme_3 As Integer
    Dim RechtsSumme_3 As Integer
    Dim SpielerSumme_3 As Integer
    Dim Links3 As Integer
    Dim Mitte3 As Integer
    Dim Rechts3 As Integer
    Dim Spieler3 As Integer
    
    Dim Dateinummer As Long
    Dim pos As Long
    Dim antwort As Long
    Dim rc As Long
    Dim Spielstand As String
    Dim msg As String
        
    Filename = App.Path & "\roserl.bin"
    If Not file_exist(Filename) Then Exit Sub
    
    Dateinummer = FreeFile
    Open Filename For Binary Access Read As #Dateinummer
    Get #Dateinummer, , strZeile1
    pos = InStr(1, strZeile1, "Spielstand vom")
    If pos = 0 Then
        Close #Dateinummer
        MsgBox "Fehler in roserl.bin. Diese Datei wird gelöscht"
        rc = file_delete(App.Path & "\roserl.bin", , True)
        End
    End If
    pos = InStr(1, strZeile1, ";")
    Spielstand = Mid(strZeile1, 1, pos)

    Spielstand = Mid(strZeile1, 1, pos - 1)
    gblnMsgBoxAntwort = True
    gstrMsg = "Willst Du mit dem " & Spielstand & " weiterspielen?"
    frmMsgBox.Show 1
    If glngAntwort = vbYes Then
        Get #Dateinummer, , LinksRot_1
        frmSpielBogen.lblLinksRot_1 = LinksRot_1
        Get #Dateinummer, , MitteRot_1
        frmSpielBogen.lblMitteRot_1 = MitteRot_1
        Get #Dateinummer, , RechtsRot_1
        frmSpielBogen.lblRechtsRot_1 = RechtsRot_1
        Get #Dateinummer, , SpielerRot_1
        frmSpielBogen.lblSpielerRot_1 = SpielerRot_1
        Get #Dateinummer, , LinksOber_1
        frmSpielBogen.lblLinksOber_1 = LinksOber_1
        Get #Dateinummer, , MitteOber_1
        frmSpielBogen.lblMItteOber_1 = MitteOber_1
        Get #Dateinummer, , RechtsOber_1
        frmSpielBogen.lblRechtsOber_1 = RechtsOber_1
        Get #Dateinummer, , SpielerOber_1
        frmSpielBogen.lblSpielerOber_1 = SpielerOber_1
        Get #Dateinummer, , LinksStiche_1
        frmSpielBogen.lblLinksStiche_1 = LinksStiche_1
        Get #Dateinummer, , MitteStiche_1
        frmSpielBogen.lblMitteStiche_1 = MitteStiche_1
        Get #Dateinummer, , RechtsStiche_1
        frmSpielBogen.lblRechtsStiche_1 = RechtsStiche_1
        Get #Dateinummer, , SpielerStiche_1
        frmSpielBogen.lblSpielerStiche_1 = SpielerStiche_1
        Get #Dateinummer, , LinksEuL_1
        frmSpielBogen.lblLinksEuL_1 = LinksEuL_1
        Get #Dateinummer, , MitteEuL_1
        frmSpielBogen.lblMitteEuL_1 = MitteEuL_1
        Get #Dateinummer, , RechtsEuL_1
        frmSpielBogen.lblRechtsEuL_1 = RechtsEuL_1
        Get #Dateinummer, , SpielerEuL_1
        frmSpielBogen.lblSpielerEuL_1 = SpielerEuL_1
        Get #Dateinummer, , LinksRK_1
        frmSpielBogen.lblLinksRK_1 = LinksRK_1
        Get #Dateinummer, , MitteRK_1
        frmSpielBogen.lblMitteRK_1 = MitteRK_1
        Get #Dateinummer, , RechtsRK_1
        frmSpielBogen.lblRechtsRK_1 = RechtsRK_1
        Get #Dateinummer, , SpielerRK_1
        frmSpielBogen.lblSpielerRK_1 = SpielerRK_1
        Get #Dateinummer, , LinksLorum_1
        frmSpielBogen.lblLinksLorum_1 = LinksLorum_1
        Get #Dateinummer, , MitteLorum_1
        frmSpielBogen.lblMitteLorum_1 = MitteLorum_1
        Get #Dateinummer, , RechtsLorum_1
        frmSpielBogen.lblRechtsLorum_1 = RechtsLorum_1
        Get #Dateinummer, , SpielerLorum_1
        frmSpielBogen.lblSpielerLorum_1 = SpielerLorum_1
        Get #Dateinummer, , LinksSumme_1
        frmSpielBogen.lblLinksSumme_1 = LinksSumme_1
        Get #Dateinummer, , MitteSumme_1
        frmSpielBogen.lblMitteSumme_1 = MitteSumme_1
        Get #Dateinummer, , RechtsSumme_1
        frmSpielBogen.lblRechtsSumme_1 = RechtsSumme_1
        Get #Dateinummer, , SpielerSumme_1
        frmSpielBogen.lblSpielerSumme_1 = SpielerSumme_1
        Get #Dateinummer, , Links1
        frmSpielBogen.lblLinksZwS_1 = Links1
        Get #Dateinummer, , Mitte1
        frmSpielBogen.lblMitteZwS_1 = Mitte1
        Get #Dateinummer, , Rechts1
        frmSpielBogen.lblRechtsZwS_1 = Rechts1
        Get #Dateinummer, , Spieler1
        frmSpielBogen.lblSpielerZwS_1 = Spieler1
        
        Get #Dateinummer, , LinksRot_2
        frmSpielBogen.lblLinksRot_2 = LinksRot_2
        Get #Dateinummer, , MitteRot_2
        frmSpielBogen.lblMitteRot_2 = MitteRot_2
        Get #Dateinummer, , RechtsRot_2
        frmSpielBogen.lblRechtsRot_2 = RechtsRot_2
        Get #Dateinummer, , SpielerRot_2
        frmSpielBogen.lblSpielerRot_2 = SpielerRot_2
        Get #Dateinummer, , LinksOber_2
        frmSpielBogen.lblLinksOber_2 = LinksOber_2
        Get #Dateinummer, , MitteOber_2
        frmSpielBogen.lblMItteOber_2 = MitteOber_2
        Get #Dateinummer, , RechtsOber_2
        frmSpielBogen.lblRechtsOber_2 = RechtsOber_2
        Get #Dateinummer, , SpielerOber_2
        frmSpielBogen.lblSpielerOber_2 = SpielerOber_2
        Get #Dateinummer, , LinksStiche_2
        frmSpielBogen.lblLinksStiche_2 = LinksStiche_2
        Get #Dateinummer, , MitteStiche_2
        frmSpielBogen.lblMitteStiche_2 = MitteStiche_2
        Get #Dateinummer, , RechtsStiche_2
        frmSpielBogen.lblRechtsStiche_2 = RechtsStiche_2
        Get #Dateinummer, , SpielerStiche_2
        frmSpielBogen.lblSpielerStiche_2 = SpielerStiche_2
        Get #Dateinummer, , LinksEuL_2
        frmSpielBogen.lblLinksEuL_2 = LinksEuL_2
        Get #Dateinummer, , MitteEuL_2
        frmSpielBogen.lblMitteEuL_2 = MitteEuL_2
        Get #Dateinummer, , RechtsEuL_2
        frmSpielBogen.lblRechtsEuL_2 = RechtsEuL_2
        Get #Dateinummer, , SpielerEuL_2
        frmSpielBogen.lblSpielerEuL_2 = SpielerEuL_2
        Get #Dateinummer, , LinksRK_2
        frmSpielBogen.lblLinksRK_2 = LinksRK_2
        Get #Dateinummer, , MitteRK_2
        frmSpielBogen.lblMitteRK_2 = MitteRK_2
        Get #Dateinummer, , RechtsRK_2
        frmSpielBogen.lblRechtsRK_2 = RechtsRK_2
        Get #Dateinummer, , SpielerRK_2
        frmSpielBogen.lblSpielerRK_2 = SpielerRK_2
        Get #Dateinummer, , LinksLorum_2
        frmSpielBogen.lblLinksLorum_2 = LinksLorum_2
        Get #Dateinummer, , MitteLorum_2
        frmSpielBogen.lblMitteLorum_2 = MitteLorum_2
        Get #Dateinummer, , RechtsLorum_2
        frmSpielBogen.lblRechtsLorum_2 = RechtsLorum_2
        Get #Dateinummer, , SpielerLorum_2
        frmSpielBogen.lblSpielerLorum_2 = SpielerLorum_2
        Get #Dateinummer, , LinksSumme_2
        frmSpielBogen.lblLinksSumme_2 = LinksSumme_2
        Get #Dateinummer, , MitteSumme_2
        frmSpielBogen.lblMitteSumme_2 = MitteSumme_2
        Get #Dateinummer, , RechtsSumme_2
        frmSpielBogen.lblRechtsSumme_2 = RechtsSumme_2
        Get #Dateinummer, , SpielerSumme_2
        frmSpielBogen.lblSpielerSumme_2 = SpielerSumme_2
        Get #Dateinummer, , Links2
        frmSpielBogen.lblLinksZwS_2 = Links2
        Get #Dateinummer, , Mitte2
        frmSpielBogen.lblMitteZwS_2 = Mitte2
        Get #Dateinummer, , Rechts2
        frmSpielBogen.lblRechtsZwS_2 = Rechts2
        Get #Dateinummer, , Spieler2
        frmSpielBogen.lblSpielerZwS_2 = Spieler2

        Get #Dateinummer, , LinksRot_3
        frmSpielBogen.lblLinksRot_3 = LinksRot_3
        Get #Dateinummer, , MitteRot_3
        frmSpielBogen.lblMitteRot_3 = MitteRot_3
        Get #Dateinummer, , RechtsRot_3
        frmSpielBogen.lblRechtsRot_3 = RechtsRot_3
        Get #Dateinummer, , SpielerRot_3
        frmSpielBogen.lblSpielerRot_3 = SpielerRot_3
        Get #Dateinummer, , LinksOber_3
        frmSpielBogen.lblLinksOber_3 = LinksOber_3
        Get #Dateinummer, , MitteOber_3
        frmSpielBogen.lblMItteOber_3 = MitteOber_3
        Get #Dateinummer, , RechtsOber_3
        frmSpielBogen.lblRechtsOber_3 = RechtsOber_3
        Get #Dateinummer, , SpielerOber_3
        frmSpielBogen.lblSpielerOber_3 = SpielerOber_3
        Get #Dateinummer, , LinksStiche_3
        frmSpielBogen.lblLinksStiche_3 = LinksStiche_3
        Get #Dateinummer, , MitteStiche_3
        frmSpielBogen.lblMitteStiche_3 = MitteStiche_3
        Get #Dateinummer, , RechtsStiche_3
        frmSpielBogen.lblRechtsStiche_3 = RechtsStiche_3
        Get #Dateinummer, , SpielerStiche_3
        frmSpielBogen.lblSpielerStiche_3 = SpielerStiche_3
        Get #Dateinummer, , LinksEuL_3
        frmSpielBogen.lblLinksEuL_3 = LinksEuL_3
        Get #Dateinummer, , MitteEuL_3
        frmSpielBogen.lblMitteEuL_3 = MitteEuL_3
        Get #Dateinummer, , RechtsEuL_3
        frmSpielBogen.lblRechtsEuL_3 = RechtsEuL_3
        Get #Dateinummer, , SpielerEuL_3
        frmSpielBogen.lblSpielerEuL_3 = SpielerEuL_3
        Get #Dateinummer, , LinksRK_3
        frmSpielBogen.lblLinksRK_3 = LinksRK_3
        Get #Dateinummer, , MitteRK_3
        frmSpielBogen.lblMitteRK_3 = MitteRK_3
        Get #Dateinummer, , RechtsRK_3
        frmSpielBogen.lblRechtsRK_3 = RechtsRK_3
        Get #Dateinummer, , SpielerRK_3
        frmSpielBogen.lblSpielerRK_3 = SpielerRK_3
        Get #Dateinummer, , LinksLorum_3
        frmSpielBogen.lblLinksLorum_3 = LinksLorum_3
        Get #Dateinummer, , MitteLorum_3
        frmSpielBogen.lblMitteLorum_3 = MitteLorum_3
        Get #Dateinummer, , RechtsLorum_3
        frmSpielBogen.lblRechtsLorum_3 = RechtsLorum_3
        Get #Dateinummer, , SpielerLorum_3
        frmSpielBogen.lblSpielerLorum_3 = SpielerLorum_3
        Get #Dateinummer, , LinksSumme_3
        frmSpielBogen.lblLinksSumme_3 = LinksSumme_3
        Get #Dateinummer, , MitteSumme_3
        frmSpielBogen.lblMitteSumme_3 = MitteSumme_3
        Get #Dateinummer, , RechtsSumme_3
        frmSpielBogen.lblRechtsSumme_3 = RechtsSumme_3
        Get #Dateinummer, , SpielerSumme_3
        frmSpielBogen.lblSpielerSumme_3 = SpielerSumme_3
        Get #Dateinummer, , Links3
        frmSpielBogen.lblLinksZwS_3 = Links3
        Get #Dateinummer, , Mitte3
        frmSpielBogen.lblMitteZwS_3 = Mitte3
        Get #Dateinummer, , Rechts3
        frmSpielBogen.lblRechtsZwS_3 = Rechts3
        Get #Dateinummer, , Spieler3
        frmSpielBogen.lblSpielerZwS_3 = Spieler3

        If frmSpielBogen.lblSpielerZwS_3 <> "0" Then
            Rundezähler = 4
            frmRoserl.lblRunde.Caption = "4. Runde"
            AusspielerBeiSpielbeginn = Spieler
            Ausspieler = Spieler
        ElseIf frmSpielBogen.lblSpielerZwS_2 <> "0" Then
            Rundezähler = 3
            frmRoserl.lblRunde.Caption = "3. Runde"
            AusspielerBeiSpielbeginn = Rechts
            Ausspieler = Rechts
        ElseIf frmSpielBogen.lblSpielerZwS_1 <> "0" Then
            Rundezähler = 2
            frmRoserl.lblRunde.Caption = "2. Runde"
            AusspielerBeiSpielbeginn = Mitte
            Ausspieler = Mitte
            AusspielerBeiSpielbeginn = Mitte
            Ausspieler = Mitte
        End If
    End If
    Close #Dateinummer
    rc = file_delete(App.Path & "\roserl.bin", , True)
End Sub
