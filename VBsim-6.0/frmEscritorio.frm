VERSION 5.00
Begin VB.Form frmEscritorio 
   Caption         =   "Escritorio"
   ClientHeight    =   6180
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11070
   Icon            =   "frmEscritorio.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmEscritorio.frx":1084A
   ScaleHeight     =   6180
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameBotones 
      BorderStyle     =   0  'None
      Caption         =   "Botones"
      Height          =   3015
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   2295
      Begin VB.CommandButton cmdMozilla 
         BackColor       =   &H000040C0&
         Caption         =   "Mozilla"
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdExplorador 
         BackColor       =   &H00404040&
         Caption         =   "Explorador"
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdTerminal 
         BackColor       =   &H00000080&
         Caption         =   "Terminal"
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   1815
      End
   End
   Begin VB.Frame frameFondo 
      Caption         =   "Fondos de Pantalla"
      Height          =   3375
      Left            =   6960
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
      Begin VB.OptionButton optFondo3 
         Caption         =   "Fondo3"
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   2040
         Width           =   3135
      End
      Begin VB.OptionButton optFondo2 
         Caption         =   "Fondo2"
         Height          =   1095
         Left            =   480
         TabIndex        =   2
         Top             =   1200
         Width           =   3015
      End
      Begin VB.OptionButton optFondo1 
         Caption         =   "Fondo1"
         Height          =   735
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "Sistema"
      Begin VB.Menu mnuSalir 
         Caption         =   "¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½"
      End
   End
   Begin VB.Menu mnuApps 
      Caption         =   "Aplicaciones"
      Begin VB.Menu mnuOffice 
         Caption         =   "¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Of"
         Begin VB.Menu mnuWord 
            Caption         =   "½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½"
         End
      End
   End
   Begin VB.Menu mnuPanelControl 
      Caption         =   "Panel de Control"
      Begin VB.Menu mnuFondos 
         Caption         =   "¿½ï¿½ï¿½ï¿½ï¿½ï¿½ï¿½Fondos"
      End
   End
End
Attribute VB_Name = "frmEscritorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExplorador_Click()
    Dim X As Integer
    X = Shell("C:\WINDOWS\system32\explorer.exe")
End Sub

Private Sub cmdMozilla_Click()
    Dim X As Integer
    X = Shell("C:\Program Files\Mozilla Firefox\firefox.exe")
End Sub

Private Sub cmdTerminal_Click()
    Dim X As Integer
    X = Shell("C:\WINDOWS\system32\cmd.exe")
End Sub

Private Sub cmdTerminal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmdExplorador.Enabled = False
    Me.cmdMozilla.Enabled = False
    Me.cmdTerminal.Enabled = True
End Sub

Private Sub cmdExplorador_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmdExplorador.Enabled = True
    Me.cmdMozilla.Enabled = False
    Me.cmdTerminal.Enabled = False
End Sub



Private Sub cmdMozilla_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmdExplorador.Enabled = False
    Me.cmdMozilla.Enabled = True
    Me.cmdTerminal.Enabled = False
End Sub


Private Sub Form_Load()
    frmEscritorio.WindowState = 2
End Sub

Private Sub frameBotones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmdExplorador.Enabled = True
    Me.cmdMozilla.Enabled = True
    Me.cmdTerminal.Enabled = True
End Sub


Private Sub mnuFondos_Click()
    If Me.frameFondo.Visible = True Then
        Me.frameFondo.Visible = False
    Else
        Me.frameFondo.Visible = True
    End If
End Sub

Private Sub mnuSalir_Click()
    'MsgBox "Adios !! Regresa pronto"
    'End
    Dim Y As Integer
    Y = MsgBox("adios", vbYesNo + vbQuestion + vbDefaultButton2, "XenomorfOS")
    frmEscritorio.Caption = Y
    If Y = 6 Then
        MsgBox "Adios !! Regresa pronto"
        End
    Else
        MsgBox "No estoy jugando"
    End If
End Sub


Private Sub optFondo1_Click()
    Me.Picture = LoadPicture("C:\Users\maye_\Downloads\alien-movie.jpg")
End Sub

Private Sub optFondo2_Click()
    Me.Picture = LoadPicture("C:\users\maye_\Pictures\Desktop\nightmare.jpg")
End Sub

Private Sub optFondo3_Click()
    Me.Picture = LoadPicture("C:\Users\maye_\Downloads\alien-movie2.jpeg")
End Sub
