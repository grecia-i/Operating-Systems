VERSION 5.00
Begin VB.Form frmSesion 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Inicio de Sesion"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   1095
      Left            =   5280
      Picture         =   "frmSesion.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   3960
      Picture         =   "frmSesion.frx":0B3B
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   6600
      Picture         =   "frmSesion.frx":123A
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdArtista 
      BackColor       =   &H00404000&
      Caption         =   "@ Cred. ID: Avvunka (twt)"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   735
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdInicio 
      Caption         =   "Iniciar Sesión"
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   240
      Picture         =   "frmSesion.frx":1D10
      ScaleHeight     =   4275
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblAutor 
      BackStyle       =   0  'Transparent
      Caption         =   "Por : @ Grecia Irais Meneses Calderas"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label lblPregunta 
      BackStyle       =   0  'Transparent
      Caption         =   "¿Qué quiere hacer?"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblIntro 
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenid@ a una simulación preliminar del Sistema Operativo XenomorfiOS"
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lblSO 
      Caption         =   "XenomorfiOS"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "frmSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Para que pueda abrir un link en navegador
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
     ByVal lpOperation As String, _
     ByVal lpFile As String, _
     ByVal lpParameters As String, _
     ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As Long


Private Sub cmdArtista_Click()
    Dim URL As String
    Dim X As Long
    
    URL = "https://x.com/Avvunka?t=wqAYn6XcppKu5IGPyr5EJQ&s=09"  'Creditos al artista
    
    X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)

    If X <= 32 Then
        MsgBox "No se pudo abrir el enlace", vbExclamation, "Error"
    End If
End Sub

Private Sub cmdInicio_Click()
    frmValidacion.Show
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    MsgBox "Adios !! Regresa pronto."
    End
End Sub

Private Sub Form_Load()
    frmEscritorio.Hide
End Sub

Private Sub lblAutor_Click()
    Dim URL As String
    Dim X As Long
    
    URL = "https://mx.pinterest.com/iKissYuna/"  'Mi perfil de Pinterest !
    X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)

    If X <= 32 Then
        MsgBox "No se pudo abrir el enlace", vbExclamation, "Error"
    End If
End Sub
