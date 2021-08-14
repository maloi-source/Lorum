VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Roserl"
   ClientHeight    =   1236
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   8892
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1236
   ScaleWidth      =   8892
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnNein 
      Caption         =   "Nein"
      Height          =   372
      Left            =   4320
      TabIndex        =   3
      Top             =   600
      Width           =   1332
   End
   Begin VB.CommandButton btnYes 
      Caption         =   "Ja"
      Height          =   372
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   1332
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label lblMsgBox 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      Caption         =   "123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-123456789-"
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7452
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim SaveWidth As Long
    Dim SaveHeight As Long
    Dim SaveTop As Long
    Dim SaveLeft As Long
    Dim blnResize As Boolean

Private Sub btnNein_Click()
    gblnMsgBoxAntwort = False
    glngAntwort = vbNo
    Unload Me
End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub btnYes_Click()
    gblnMsgBoxAntwort = False
    glngAntwort = vbYes
    Unload Me
End Sub

Private Sub Form_Load()
    'Die Anzahl Zeilen in gstrMsg wird gezählt - Multiplizieren mit dem Wert für klein/mittel/gross=lblMsgBox.Width
    'Die breiteste Zeile in gstrMsg ermitteln - Multiplizieren mit dem Wert für klein/mittel/gross=lblMsgBox.Height
    'Die Form in Fenstermitte positionieren - Form.Height \ 2 und Form.Width \ 2
    '
    Dim AnzahlZeilen As Long
    Dim BreitesteZeile As Long
    Dim pos As Long
    Dim start As Long
    Dim PixelKleinMittelGross As Long
    Dim Zeilenhöhe As Long
    Dim msg As String
    
    If gstrMsg = "" Then
        Unload Me
    End If
    
    If PublicCheckForDPI = "1" Then
        PixelKleinMittelGross = 8 * Screen.TwipsPerPixelX
        Zeilenhöhe = 20 * Screen.TwipsPerPixelY
    ElseIf PublicCheckForDPI = "2" Then
        PixelKleinMittelGross = 10 * Screen.TwipsPerPixelX
        Zeilenhöhe = 25 * Screen.TwipsPerPixelY
    ElseIf PublicCheckForDPI = "3" Then
        PixelKleinMittelGross = 12 * Screen.TwipsPerPixelX
        Zeilenhöhe = 30 * Screen.TwipsPerPixelY
    End If
    start = 1
    AnzahlZeilen = 1
    BreitesteZeile = 25
    Do
        pos = InStr(start, gstrMsg, vbNewLine)
        If pos = 0 Then Exit Do
        If pos - start > BreitesteZeile Then
            BreitesteZeile = pos - start
        End If
        AnzahlZeilen = AnzahlZeilen + 1
        start = pos + 1
    Loop
    If AnzahlZeilen = 1 Then
        BreitesteZeile = Len(gstrMsg)
    End If
    lblMsgBox = gstrMsg
    gstrMsg = ""
    lblMsgBox.Width = BreitesteZeile * PixelKleinMittelGross
    lblMsgBox.Height = AnzahlZeilen * Zeilenhöhe
    If gblnMsgBoxAntwort = True Then
        btnYes.Top = lblMsgBox.Height + 200
        btnNein.Top = lblMsgBox.Height + 200
        btnYes.Left = (lblMsgBox.Width \ 2) - btnYes.Width - 100
        btnNein.Left = (lblMsgBox.Width \ 2) + 100
        btnOK.Top = lblMsgBox.Height + 200                                      'Gerbing 17.08.2017 wird gebraucht für frmMsgbox.Height
        btnOK.Visible = False
    Else
        btnOK.Top = lblMsgBox.Height + 200
        btnOK.Left = (lblMsgBox.Width \ 2) - btnOK.Width \ 2
        btnYes.Visible = False
        btnNein.Visible = False
    End If
    blnResize = False
'-----------------------------------------------------------------------------------------------
'    If Mons.Count(True) <> 1 Then                                          'Gerbing 11.09.2017
'        msg = "Mons.Count=" & Mons.Count(True) & vbNewLine
'        msg = msg & "PrimaryMon=" & Mons.PrimaryMon & vbNewLine
'        msg = msg & "Mons.WorkTop(1)=" & Mons.WorkTop(1) & vbNewLine
'        msg = msg & "Mons.WorkLeft(1)=" & Mons.WorkLeft(1) & vbNewLine
'        msg = msg & "Mons.WorkWidth(1)=" & Mons.WorkWidth(1) & vbNewLine
'        msg = msg & "Mons.WorkHeight(1)=" & Mons.WorkHeight(1) & vbNewLine
'        If Mons.Enabled(2) = True Then
'            msg = msg & "Mons.WorkTop(2)=" & Mons.WorkTop(2) & vbNewLine
'            msg = msg & "Mons.WorkLeft(2)=" & Mons.WorkLeft(2) & vbNewLine
'            msg = msg & "Mons.WorkWidth(2)=" & Mons.WorkWidth(2) & vbNewLine
'            msg = msg & "Mons.WorkHeight(2)=" & Mons.WorkHeight(2) & vbNewLine
'        End If
'        If Mons.Enabled(3) = True Then
'            msg = msg & "Mons.WorkTop(3)=" & Mons.WorkTop(3) & vbNewLine
'            msg = msg & "Mons.WorkLeft(3)=" & Mons.WorkLeft(3) & vbNewLine
'            msg = msg & "Mons.WorkWidth(3)=" & Mons.WorkWidth(3) & vbNewLine
'            msg = msg & "Mons.WorkHeight(3)=" & Mons.WorkHeight(3) & vbNewLine
'        End If
'        If Mons.WorkWidth(1) = frmRoserl.ScaleWidth Then
'            msg = msg & "frmRoserl läuft auf PrimaryMon"
'        Else
'            msg = msg & "frmRoserl läuft auf ZweitMon"
'        End If
'        MsgBox msg
'    End If
    
'    If Mons.Count(True) <> 1 Then
'        If Mons.WorkWidth(1) <> frmRoserl.ScaleWidth Then
'            'frmRoserl läuft auf ZweitMon
'            If Mons.Enabled(2) = True Then
'                frmMsgBox.Width = lblMsgBox.Width + 20
'                frmMsgBox.Height = btnOK.Top + btnOK.Height + 600
'                frmMsgBox.Left = (Mons.WorkLeft(2) \ 2) - frmMsgBox.Width \ 2
'                frmMsgBox.Top = (Mons.WorkTop(2) \ 2) - frmMsgBox.Height \ 2
'            End If
'            If Mons.Enabled(3) = True Then
'                frmMsgBox.Width = lblMsgBox.Width + 20
'                frmMsgBox.Height = btnOK.Top + btnOK.Height + 600
'                frmMsgBox.Left = ((Mons.WorkLeft(3) * Screen.TwipsPerPixelX) \ 2) - frmMsgBox.Width \ 2
'                frmMsgBox.Top = ((Mons.WorkTop(3) * Screen.TwipsPerPixelY) \ 2) - frmMsgBox.Height \ 2
'                'frmMsgBox.Top = frmMsgBox.Top + 5000          'Gerbing 11.09.2017 Durch Test ermittelte Korrektur, sonst ist frmMsgbox.Top zu weit oben
'            End If
'        Else
'            frmMsgBox.Width = lblMsgBox.Width + 20
'            frmMsgBox.Height = btnOK.Top + btnOK.Height + 600
'            frmMsgBox.Left = ((screenWidth * Screen.TwipsPerPixelX) \ 2) - frmMsgBox.Width \ 2
'            frmMsgBox.Top = ((screenHeight * Screen.TwipsPerPixelY) \ 2) - frmMsgBox.Height \ 2
'        End If
'    Else
        frmMsgBox.Width = lblMsgBox.Width + 20
        frmMsgBox.Height = btnOK.Top + btnOK.Height + 600
        frmMsgBox.Left = ((screenWidth * Screen.TwipsPerPixelX) \ 2) - frmMsgBox.Width \ 2
        frmMsgBox.Top = ((screenHeight * Screen.TwipsPerPixelY) \ 2) - frmMsgBox.Height \ 2
'    End If
'---------------------------------------------------------------------'Gerbing 11.09.2017
    Call AnpassenNutzerWunsch(Me)
'    SaveWidth = frmMsgBox.Width
'    SaveHeight = frmMsgBox.Height
'    SaveTop = frmMsgBox.Top
'    SaveLeft = frmMsgBox.Left
'    blnResize = True
'    If InStr(1, lblMsgBox, "zugeben", vbTextCompare) <> 0 Then
'        msg = "frmMsgBox.Left = " & frmMsgBox.Left & vbNewLine
'        msg = msg & "frmMsgBox.Top = " & frmMsgBox.Top & vbNewLine
'        MsgBox msg
'    End If
End Sub
'
'Private Sub Form_Resize()
'    If blnResize = False Then Exit Sub
'    frmMsgBox.Width = SaveWidth
'    frmMsgBox.Height = SaveHeight
'    frmMsgBox.Top = SaveTop
'    frmMsgBox.Left = SaveLeft
'End Sub

