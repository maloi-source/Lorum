VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBestwerteZeigen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Bestwerte zeigen"
   ClientHeight    =   6948
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6672
   Icon            =   "frmBestwerteZeigen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6948
   ScaleWidth      =   6672
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btnLöschen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bestwerte löschen"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6372
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6132
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6492
      _ExtentX        =   11451
      _ExtentY        =   10816
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollBars      =   2
   End
End
Attribute VB_Name = "frmBestwerteZeigen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim LastSortCol As Integer

Private Sub btnLöschen_Click()
    On Error Resume Next
    LogFso.DeleteFile App.Path & "\RoserlBestwerte.log"                 'Gerbing 14.11.2017
    On Error GoTo 0
    Call Form_Load
End Sub

Private Sub Form_Load()                                                 'Gerbing 13.09.2017
    Dim strRoserlLog As String
    
    Call AnpassenNutzerWunsch(Me)
    Grid1.Cols = 3
    Grid1.Rows = 1
    Grid1.ColAlignment(0) = flexAlignCenterCenter
    Grid1.ColAlignment(1) = flexAlignCenterCenter
    Grid1.ColAlignment(2) = flexAlignCenterCenter
    Grid1.ColWidth(0) = (Grid1.Width - 100) / 2
    Grid1.ColWidth(1) = (Grid1.Width - 100) / 4
    Grid1.ColWidth(2) = (Grid1.Width - 100) / 4
    Grid1.TextMatrix(0, 0) = "Datum"
    Grid1.TextMatrix(0, 1) = "Spieler"
    Grid1.TextMatrix(0, 2) = "Punkte"
    
    'object.OpenFile(filename[, ForReading[, TristateFalse]])
    On Error GoTo NoFileFound
    
    Set oStream = LogFso.OpenTextFile(App.Path & "\RoserlBestwerte.log", ForReading, TristateFalse)
    'Datei Roserl.log Zeile für Zeile einlesen und diese Zeile in Grid1 eintragen
    Do Until oStream.AtEndOfStream = True
        strRoserlLog = oStream.ReadLine
        Grid1.AddItem strRoserlLog
    Loop
    oStream.Close
NoFileFound:
End Sub

Private Sub Grid1_Click()
    On Error GoTo Grid1Error
    With Grid1
        .Redraw = False
        If .MouseRow = 0 And .MouseCol > .FixedCols - 1 Then
            .Col = .MouseCol
            If LastSortCol <> .Col Then .Tag = 0
            If .Tag = 0 Or .Tag = 6 Then
                .Sort = 5
                .Tag = 5
                LastSortCol = .Col
            Else
                .Sort = 6
                .Tag = 6
                LastSortCol = .Col
            End If
            .Col = 0
            .Row = 0
        End If
        .Redraw = True
    End With
Grid1Error:
End Sub

