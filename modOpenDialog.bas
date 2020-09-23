Attribute VB_Name = "modOpenDialog"
Option Explicit

Private Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_FILEMUSTEXIST = &H1000

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Function OpenDialog(hwnd As Long, Filter As String, InitDir As String) As String
    
    Dim ofn     As OPENFILENAME
    Dim RetVal  As Long
    Dim idx     As Long
    
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hwnd
        .hInstance = App.hInstance
        If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
        For idx = 1 To Len(Filter)
            If Mid$(Filter, idx, 1) = "|" Then Mid$(Filter, idx, 1) = Chr$(0)
        Next
        .lpstrFilter = Filter
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = InitDir
        .lpstrTitle = "Open Picture - " & Filter
        .flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        RetVal = GetOpenFileName(ofn)
        If (RetVal) Then
            OpenDialog = Trim$(.lpstrFile)
        Else
            OpenDialog = ""
        End If
    End With
    
End Function
