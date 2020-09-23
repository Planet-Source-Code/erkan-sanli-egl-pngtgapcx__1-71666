VERSION 5.00
Begin VB.Form frmCanvas 
   ClientHeight    =   8085
   ClientLeft      =   1440
   ClientTop       =   1965
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   12480
   Begin VB.CommandButton Command1 
      Caption         =   "PNG"
      Height          =   372
      Index           =   0
      Left            =   7080
      TabIndex        =   3
      Top             =   7560
      Width           =   1572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PCX"
      Height          =   372
      Index           =   2
      Left            =   10440
      TabIndex        =   2
      Top             =   7560
      Width           =   1572
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7365
      Left            =   0
      Picture         =   "frmCanvas.frx":0000
      ScaleHeight     =   491
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   832
      TabIndex        =   1
      Top             =   0
      Width           =   12480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TGA"
      Height          =   372
      Index           =   1
      Left            =   8760
      TabIndex        =   0
      Top             =   7560
      Width           =   1572
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tga As clsLoadTGA
Dim pcx As clsLoadPCX
Dim png As clsLoadPNG

Private Sub Command1_Click(Index As Integer)
        
    Dim Filename As String
    
    mPicBox.Picture = LoadPicture(App.Path & "\desen.jpg")
    Select Case Index
        Case 0
            Filename = OpenDialog(frmCanvas.hwnd, "*.png |*.png", App.Path & "\PNG")
            If Len(Filename) Then
                Set png = New clsLoadPNG
                Call png.LoadPNG(Filename)
                Set png = Nothing
            End If
        Case 1
            Filename = OpenDialog(frmCanvas.hwnd, "*.tga |*.tga", App.Path & "\TGA")
            If Len(Filename) Then
                Set tga = New clsLoadTGA
                Call tga.LoadTGA(Filename)
                Set tga = Nothing
            End If
        Case 2
            Filename = OpenDialog(frmCanvas.hwnd, "*.pcx |*.pcx", App.Path & "\PCX")
            If Len(Filename) Then
                Set pcx = New clsLoadPCX
                Call pcx.LoadPCX(Filename)
                Set pcx = Nothing
            End If
    End Select
    
End Sub

Private Sub Form_Load()
    
    mAlpha = True
    Set mPicBox = Me.picCanvas

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set mPicBox = Nothing
    End

End Sub
