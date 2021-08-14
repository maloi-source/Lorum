Attribute VB_Name = "Module1"
Option Explicit
    Public FarbeAusspieler
    Public gblnMsgBoxAntwort As Boolean
    Public glngAntwort As Long
    Public gstrMsg As String
    Public Sprache As Long
    Dim RoserlIniFile As String
    Public Declare Function GetPrivateProfileStringW Lib "kernel32.dll" _
            (ByVal lpApplicationName As Any, _
            ByVal lpKeyName As Any, _
            ByVal lpDefault As Any, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long
    
    Public Declare Function WritePrivateProfileStringW Lib "kernel32.dll" _
            (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpString As Long, _
            ByVal lpFileName As Long) As Long
        
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hWnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
    
    Public LogFso As Scripting.FileSystemObject                                     'Gerbing 13.09.2017
    Public IniFso As Scripting.FileSystemObject
    Public oStream As Scripting.TextStream
    Dim ABSCHNITT As String * 300            'was nicht reinpaßt wird abgeschnitten
    Dim absch As String
    Dim Language As String
    Dim zeichen As Integer
    Dim StartPos As Integer
    Dim Zeile As String
    Public PublicCheckForDPI As String
    Dim CheckForDPI As String
    Public PublicLanguage As String
    
Sub ReadRoserlIniFile()
    RoserlIniFile = App.Path & "\Roserl.ini"      'Pfad der fotos.ini                 Gerbing 17.02.2011
    On Error Resume Next
    If file_path_exist(RoserlIniFile) = False Then
    'If Kontrolle = "" Then
        'Standardwerte eintragen, wenn es keine Datei RoserlINiFile.ini gibt    'Gerbing 27.09.2010
        PublicCheckForDPI = "1"                 'Gerbing 23.06.2011
        Exit Sub
    End If
    Call GlL                                'Prüfe [Global] Language
    Call DPI                                'Prüfe [Adjustments] CheckForDPI                'Gerbing 23.06.2011
End Sub

Public Function INIReadString(ByVal Section As String, ByVal key As String, ByVal Default As String, ByVal Filename As String) As String
    Dim cSize As Long
    Dim strReturn As String
    Dim retVal As Long
    
    cSize = 300
    strReturn = String(cSize, 0)
    retVal = GetPrivateProfileStringW(StrPtr(Section), StrPtr(key), StrPtr(Default), StrPtr(strReturn), cSize, StrPtr(Filename))
    If retVal > 0 Then
        INIReadString = Left(strReturn, retVal)
    End If
End Function

Sub WriteGlL(NeuerInhalt As String)
    'Schreibe [Global] Language
    If file_path_exist(RoserlIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(RoserlIniFile, True, True)
        oStream.Close
    End If
    absch = "Global"
    Language = "Language"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(Language), StrPtr(NeuerInhalt), StrPtr(RoserlIniFile))
End Sub

Sub WriteDPI(NeuerInhalt As String)                                                        'Gerbing 23.06.2011
    'Schreibe [Adjustments] CheckForDPI
    If file_path_exist(RoserlIniFile) = False Then
'        'Wenn fotos.ini nicht existiert, wird sie erzeugt
        'object.CreateTextFile(filename[, overwrite[, unicode]])
        Set oStream = IniFso.CreateTextFile(RoserlIniFile, True, True)
        oStream.Close
    End If
    absch = "Adjustments"
    CheckForDPI = "CheckForDPI"
    zeichen = WritePrivateProfileStringW(StrPtr(absch), StrPtr(CheckForDPI), StrPtr(NeuerInhalt), StrPtr(RoserlIniFile))
End Sub

Sub GlL()
    'Prüfe [Global] Language
    absch = "Global"
    Language = "Language"
    Zeile = INIReadString(absch, Language, "", RoserlIniFile)
    If Zeile = "" Then Sprache = 0            'Deutsch
    If Zeile = "0" Then Sprache = 0            'Deutsch
    If Zeile = "1" Then Sprache = 3000         'English
    'If Dir(AppPath & "\fotos.mdb") <> "" Then
    PublicLanguage = Zeile
End Sub

Sub DPI()                                                                                       'Gerbing 11.03.2017
    'Prüfe [Adjustments] CheckForDPI
    absch = "Adjustments"
    CheckForDPI = "CheckForDPI"
    Zeile = INIReadString(absch, CheckForDPI, "", RoserlIniFile)
    If Zeile = "" Then PublicCheckForDPI = "1"              'leer berücksichtigen
    If Zeile = "1" Then PublicCheckForDPI = "1"             'klein
    If Zeile = "2" Then PublicCheckForDPI = "2"             'mittel
    If Zeile = "3" Then PublicCheckForDPI = "3"             'gross
End Sub

Public Function GetShellError(lErrorCode As Long) As String
    Const SE_ERR_FNF = 2&, SE_ERR_PNF = 3&
    Const SE_ERR_ACCESSDENIED = 5&, SE_ERR_OOM = 8&
    Const SE_ERR_DLLNOTFOUND = 32&, SE_ERR_SHARE = 26&
    Const SE_ERR_ASSOCINCOMPLETE = 27&, SE_ERR_DDETIMEOUT = 28&
    Const SE_ERR_DDEFAIL = 29&, SE_ERR_DDEBUSY = 30&
    Const SE_ERR_NOASSOC = 31&, ERROR_BAD_FORMAT = 11&

    Select Case lErrorCode
        Case SE_ERR_FNF
            GetShellError = "File not found"
        Case SE_ERR_PNF
            GetShellError = "Path not found"
        Case SE_ERR_ACCESSDENIED
            GetShellError = "Access denied"
        Case SE_ERR_OOM
            GetShellError = "Out of memory"
        Case SE_ERR_DLLNOTFOUND
            GetShellError = "DLL not found"
        Case SE_ERR_SHARE
            GetShellError = "A sharing violation occurred"
        Case SE_ERR_ASSOCINCOMPLETE
            GetShellError = "Incomplete or invalid file association"
        Case SE_ERR_DDETIMEOUT
            GetShellError = "DDE Time out"
        Case SE_ERR_DDEFAIL
            GetShellError = "DDE transaction failed"
        Case SE_ERR_DDEBUSY
            GetShellError = "DDE busy"
        Case SE_ERR_NOASSOC
            GetShellError = "No association for file extension"
        Case ERROR_BAD_FORMAT
            GetShellError = "Invalid EXE file or error in EXE image"
        Case Else
            GetShellError = "Unknown error"
    End Select
End Function

