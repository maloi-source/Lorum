VERSION 5.00
Begin VB.UserControl Roserl 
   BackColor       =   &H00008000&
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1068
   ClipControls    =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   1068
   Begin VB.Shape shpSelected 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Strich
      BorderWidth     =   5
      Height          =   1440
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "Roserl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Public Enum CardTypes
        Schell = 1
        Rot
        Grün
        Eichel
    End Enum
    
    Public Enum CardValues
        Sieben = 1
        Acht
        Neun
        Unter
        Ober
        König
        Zehn
        Ass
    End Enum
    
    Public Enum FaceModes
        NachOben
        Rückseite
        Base
        Circled
        Crossed
        Back
    End Enum
    
    Public Enum GespieltVonValues
        Spieler = 1
        Links
        Mitte
        Rechts
    End Enum
    
    Public Enum EingesammeltVonValues
        Spieler = 1
        Links
        Mite
        Rechts
    End Enum

    
    Dim cValue As CardValues
    Dim cType As CardTypes
    Dim cFaceMode As FaceModes
    
    Dim cSelected As Boolean
    Dim BackPic As Integer
    
    Dim cGespieltVon As GespieltVonValues
    Dim cEingesammeltVon As EingesammeltVonValues
    
    Event SelectCard()
    Event DeselectCard()
    Event Click()
    Event DblClick()
    

' ----==== GDI+ Konstanten ====----
Private Const GdiPlusVersion As Long = 1&

' ----==== GDI+ Typen ====----
Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type

' ----==== Sonstige Typen ====----
Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

' ----==== GDI+ Enums ====----
Private Enum Status 'GDI+ Status
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

' ----==== GDI+ API Deklarationen ====----
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
    (ByVal bitmap As Long, ByRef hbmReturn As Long, _
    ByVal Background As Long) As Status

Private Declare Function GdipDisposeImage Lib "gdiplus" _
    (ByVal image As Long) As Status

Private Declare Function GdipLoadImageFromStream Lib "gdiplus" _
    (ByVal Stream As Any, ByRef image As Long) As Status

Private Declare Function GdiplusShutdown Lib "gdiplus" _
    (ByVal token As Long) As Status

Private Declare Function GdiplusStartup Lib "gdiplus" _
    (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, _
    ByRef lpOutput As GdiplusStartupOutput) As Status

' ----==== OLE32 API Deklarationen ====----
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" _
    (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, _
    ByRef ppstm As Any)

' ----==== OLEAUT32 API Deklarations ====----
Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" _
    (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, _
    lplpvObj As Object)

' ----==== Variablen ====----
Dim GdipToken As Long
Dim GdipInitialized As Boolean


        
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
    Private StartMillisek As Long
    Private EndMillisek As Long

Property Get FaceMode() As FaceModes
    FaceMode = cFaceMode
End Property

Property Let FaceMode(NewValue As FaceModes)
    cFaceMode = NewValue
    RefreshCard
    PropertyChanged "FaceMode"
End Property

Property Get GespieltVon() As GespieltVonValues
    GespieltVon = cGespieltVon
End Property

Property Let GespieltVon(NewValue As GespieltVonValues)
    cGespieltVon = NewValue
    RefreshCard
    PropertyChanged "GespieltVon"
End Property

Property Get EingesammeltVon() As EingesammeltVonValues
    EingesammeltVon = cEingesammeltVon
End Property

Property Let EingesammeltVon(NewValue As EingesammeltVonValues)
    cEingesammeltVon = NewValue
    RefreshCard
    PropertyChanged "EingesammeltVon"
End Property

Property Get Selected() As Boolean
    Selected = cSelected
End Property

Property Let Selected(NewValue As Boolean)
    If cSelected = NewValue Then Exit Property
    cSelected = NewValue
    If NewValue Then RaiseEvent SelectCard Else: RaiseEvent DeselectCard
    RefreshCard
    PropertyChanged "Selected"
End Property

Sub RefreshCard()
    shpSelected.Visible = cSelected
    If cFaceMode = NachOben Then
        GetCard
    Else
        'FaceDown
        GetFace
    End If
End Sub

Property Get CardValue() As CardValues
    CardValue = cValue
End Property

Property Let CardValue(NewValue As CardValues)
    cValue = NewValue
    RefreshCard
    PropertyChanged "CardValue"
End Property

Property Get CardType() As CardTypes
    CardType = cType
End Property

Property Let CardType(NewValue As CardTypes)
    cType = NewValue
    RefreshCard
    PropertyChanged "CardType"
End Property

Private Sub imgCard_Click()
    RaiseEvent Click
End Sub

Private Sub imgCard_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    cValue = Ass
    cType = Grün
    UserControl.Width = 1065
    UserControl.Height = 1440
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    cFaceMode = PropBag.ReadProperty("FaceMode", Rückseite)
    cSelected = PropBag.ReadProperty("Selected", False)
    cValue = PropBag.ReadProperty("CardValue", Ass)
    'cType = PropBag.ReadProperty("CardType", Grün)
    cType = PropBag.ReadProperty("CardType", Schell)
    RefreshCard
End Sub

Private Sub UserControl_Resize()
    imgCard.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Sub GetCard()
    If cValue = 0 Then Exit Sub
    'Set imgCard.Picture = LoadResPicture(100 + 4 * (cValue - 1) + cType, 0)
    'Set imgCard.Picture = LoadImageFromRES(100 + 4 * (cValue - 1) + cType, "GIF")
    'Set imgCard.Picture = LoadImageFromRES(100 * cType + cValue, "GIF")
    'Set imgCard.Picture = LoadImageFromRES(100 * cType + cValue, "GIF")
    StartMillisek = timeGetTime
    Set imgCard.Picture = LoadImageFromCustomRes(100 * cType + cValue, "GIF")
    EndMillisek = timeGetTime
    Debug.Print "EndMillisec=" & EndMillisek
    Debug.Print "Millisekunden für LoadImageFromCustomRes" & "=" & (EndMillisek - StartMillisek)
End Sub

Public Sub SetCardValue(cdType As CardTypes, cdValue As CardValues)
    cType = cdType
    cValue = cdValue
    RefreshCard
End Sub

Function GetFace()
    If cFaceMode = NachOben Then Exit Function
    'If cFaceMode = Rückseite Then Exit Function             'Gerbing 19.05.2017
    If cFaceMode = Back Then cFaceMode = Back + BackPic
    'Set imgCard.Picture = LoadImageFromRES(152 + cFaceMode, "GIF")
    'Set imgCard.Picture = LoadImageFromRES(153, "GIF")
    Set imgCard.Picture = LoadImageFromCustomRes(153, "GIF")
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "FaceMode", cFaceMode, Rückseite
    PropBag.WriteProperty "Selected", cSelected, False
    PropBag.WriteProperty "CardValue", cValue, Ass
    'PropBag.WriteProperty "CardType", cType, Grün
    PropBag.WriteProperty "CardType", cType, Schell
End Sub

'------------------------------------------------------
' Funktion     : Execute
' Beschreibung : Gibt im Fehlerfall die entsprechende
'                GDI+ Fehlermeldung aus
' Übergabewert : GDI+ Status
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Private Function Execute(ByVal lReturn As Status) As Status
    Dim lCurErr As Status
    If lReturn = Status.OK Then
        lCurErr = Status.OK
    Else
        lCurErr = lReturn
        MsgBox GdiErrorString(lReturn) & " GDI+ Error:" & _
        lReturn, vbOKOnly, "GDI Error"
    End If
    Execute = lCurErr
End Function

'------------------------------------------------------
' Funktion     : GdiErrorString
' Beschreibung : Umwandlung der GDI+ Statuscodes in Stringcodes
' Übergabewert : GDI+ Status
' Rückgabewert : Fehlercode als String
'------------------------------------------------------
Private Function GdiErrorString(ByVal lError As Status) As String
    Dim s As String
    
    Select Case lError
    Case GenericError:              s = "Generic Error."
    Case InvalidParameter:          s = "Invalid Parameter."
    Case OutOfMemory:               s = "Out Of Memory."
    Case ObjectBusy:                s = "Object Busy."
    Case InsufficientBuffer:        s = "Insufficient Buffer."
    Case NotImplemented:            s = "Not Implemented."
    Case Win32Error:                s = "Win32 Error."
    Case WrongState:                s = "Wrong State."
    Case Aborted:                   s = "Aborted."
    Case FileNotFound:              s = "File Not Found."
    Case ValueOverflow:             s = "Value Overflow."
    Case AccessDenied:              s = "Access Denied."
    Case UnknownImageFormat:        s = "Unknown Image Format."
    Case FontFamilyNotFound:        s = "FontFamily Not Found."
    Case FontStyleNotFound:         s = "FontStyle Not Found."
    Case NotTrueTypeFont:           s = "Not TrueType Font."
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version."
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized."
    Case PropertyNotFound:          s = "Property Not Found."
    Case PropertyNotSupported:      s = "Property Not Supported."
    Case Else:                      s = "Unknown GDI+ Error."
    End Select
    
    GdiErrorString = s
End Function

'------------------------------------------------------
' Funktion     : HandleToPicture
' Beschreibung : Umwandeln einer Bitmap Handle in
'                ein StdPicture Objekt
' Übergabewert : hGDIHandle = Bitmap Handle
'                ObjectType = Bitmaptyp
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Private Function HandleToPicture(ByVal hGDIHandle As Long, _
    ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hpal As Long = 0) As StdPicture
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As IID
    Dim oPicture As IPicture
    
    ' Initialisiert die PICTDESC Structur
    With tPictDesc
        .cbSizeOfStruct = Len(tPictDesc)
        .picType = ObjectType
        .hgdiObj = hGDIHandle
        .hPalOrXYExt = hpal
    End With
    
    ' Initialisiert das IPicture Interface ID
    With IID_IPicture
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    ' Erzeugen des Objekts
    OleCreatePictureIndirect tPictDesc, _
    IID_IPicture, True, oPicture
    
    ' Rückgabe des Pictureobjekts
    Set HandleToPicture = oPicture
    
End Function

'------------------------------------------------------
' Funktion     : LoadImageFromCustomRes
' Beschreibung : Lädt ein Bild aus einer "CUSTOM"-Ressource
'                (alle GDI+ Bildformate)
'                BMP; DIB; RLE; JPG; JPEG; JPE; JFIF; GIF
'                EMF; WMF; TIF; TIFF; PNG; ICO
' Übergabewert : ResIndex = Kennung (ID) der Daten
'                           in der Ressourcedatei
'                ResName = Zeichenfolgenname der
'                          benutzerdefinierten Ressource
' Rückgabewert : StdPicture Objekt
'------------------------------------------------------
Private Function LoadImageFromCustomRes( _
    ByVal ResIndex As Long, _
    ByVal ResName As String) As StdPicture
    
    GdipInitialized = False
    
    ' GDI+ starten
    If Execute(StartUpGDIPlus(GdiPlusVersion)) = OK Then
        GdipInitialized = True
    Else
        MsgBox "GDI+ not inizialized.", _
        vbOKOnly, "GDI Error"
    End If

    
    
    On Error GoTo PROC_ERR
    
    Dim ResData() As Byte
    Dim Stream As IUnknown
    Dim lBitmap As Long
    Dim hBitmap As Long
    
    ' Ressource in ByteArray speichern
    ResData = LoadResData(ResIndex, ResName)
    
    ' Stream erzeugen
    Call CreateStreamOnHGlobal(ResData(0), _
    False, Stream)
    
    ' ist ein Stream vorhanden
    If Not (Stream Is Nothing) Then
        
        ' GDI+ Bitmapobjekt vom Stream erstellen
        If Execute(GdipLoadImageFromStream( _
        Stream, lBitmap)) = OK Then
            
            ' Handle des Bitmapobjektes ermitteln
            If Execute(GdipCreateHBITMAPFromBitmap( _
            lBitmap, hBitmap, 0)) = OK Then
                
                ' StdPicture Objekt erstellen
                Set LoadImageFromCustomRes = _
                HandleToPicture(hBitmap, vbPicTypeBitmap)
                
            End If
            
            ' Bitmapobjekt löschen
            Call Execute(GdipDisposeImage(lBitmap))
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LoadImageFromCustomRes"
    Resume PROC_EXIT

End Function

'------------------------------------------------------
' Funktion     : ShutdownGDIPlus
' Beschreibung : Beendet die GDI+ Instanz
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Private Function ShutdownGDIPlus() As Status
    ' Beendet GDI+ Instanz
    ShutdownGDIPlus = GdiplusShutdown(GdipToken)
End Function

'------------------------------------------------------
' Funktion     : StartUpGDIPlus
' Beschreibung : Initialisiert GDI+ Instanz
' Übergabewert : GDI+ Version
' Rückgabewert : GDI+ Status
'------------------------------------------------------
Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As Status
    ' Initialisieren der GDI+ Instanz
    Dim GdipStartupInput As GDIPlusStartupInput
    Dim GdipStartupOutput As GdiplusStartupOutput
    
    GdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(GdipToken, _
    GdipStartupInput, GdipStartupOutput)
End Function




