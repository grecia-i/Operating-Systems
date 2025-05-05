VERSION 5.00
Begin VB.Form frmValidacion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmValidacion"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7560
      Top             =   4920
   End
   Begin VB.PictureBox Picture6 
      Height          =   375
      Left            =   7440
      Picture         =   "frmValidacion.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox Picture5 
      Height          =   375
      Left            =   7440
      Picture         =   "frmValidacion.frx":0378
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   3960
      TabIndex        =   11
      Text            =   "UsuarioEjemplo"
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtContra 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   10
      Text            =   "xenomorfiOS"
      Top             =   2160
      Width           =   3375
   End
   Begin VB.CommandButton cmdInicio 
      Caption         =   "Iniciar Sesión"
      Height          =   735
      Left            =   4200
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   735
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   6600
      Picture         =   "frmValidacion.frx":173B
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   3960
      Picture         =   "frmValidacion.frx":2211
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      Height          =   1095
      Left            =   5280
      Picture         =   "frmValidacion.frx":2910
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdArtista 
      BackColor       =   &H00404000&
      Caption         =   "@ Cred. ID: Avvunka (twt)"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   4800
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   240
      Picture         =   "frmValidacion.frx":344B
      ScaleHeight     =   4275
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblTiempo 
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   4920
      Width           =   1695
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
      Left            =   6120
      TabIndex        =   14
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese una Contraseña (Usuario Nuevo: Xeno)"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label lblIntro 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese un Usuario (Usuario Nuevo: Alien)"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1320
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
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "frmValidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
     ByVal lpOperation As String, _
     ByVal lpFile As String, _
     ByVal lpParameters As String, _
     ByVal lpDirectory As String, _
     ByVal nShowCmd As Long) As Long

Dim Tries As Integer
Dim Tiempo As Integer
Dim Cooldown As Boolean

Private Sub cmdInicio_Click()
    Dim user As String
    Dim contra As String
    
    user = "Alien"
    contra = "Xeno"
    Tries = Tries - 1
    
    If Cooldown Then
        MsgBox "Espere " & Tiempo & " segundos antes de intentar nuevamente.", vbExclamation, "En cooldown"
        Exit Sub
    End If
    
    If txtContra.Text = contra Then
        If txtUsuario.Text = user Then
            frmEscritorio.Show
            Unload Me
        Else
            If Tries < 1 Then
                MsgBox "Fuera de intentos. Inténtelo de nuevo cuando acabe el Cooldown", vbCritical, "Error"
                lblTiempo.Visible = True
                Tiempo = 50
                Cooldown = True
                cmdInicio.Enabled = False
                Timer1.Enabled = True
            Else
                MsgBox "Usuario incorrecto. Quedan " & (Tries) & " intentos restantes.", vbExclamation, "Error"
            End If
        End If
    
    Else
        If Tries < 1 Then
            MsgBox "Fuera de intentos. Inténtelo de nuevo cuando acabe el Cooldown", vbCritical, "Error"
            lblTiempo.Visible = True
            Tiempo = 50
            Cooldown = True
            cmdInicio.Enabled = False
            Timer1.Enabled = True
        Else
            MsgBox "Contraseña incorrecta. Quedan " & (Tries) & " intentos restantes.", vbExclamation, "Error"
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    MsgBox "Adios !! Regresa pronto."
    End
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

Private Sub lblAutor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAutor.Enabled = True
    lblAutor.ForeColor = vbBlack
    lblAutor.MousePointer = vbHandPoint
End Sub

Private Sub lblAutor_MouseLeave()
    lblAutor.ForeColor = vbBlue
    lblAutor.Font.Underline = False
    lblAutor.MousePointer = vbDefault
    End With
End Sub


Private Sub Picture6_Click()
    If txtContra.PasswordChar = "*" Then
        txtContra.PasswordChar = ""
    Else
        txtContra.PasswordChar = "*"
    End If
End Sub

Private Sub Form_Load()
    frmSesion.Hide
    frmEscritorio.Hide
    Tries = 5
    Cooldown = False
    Timer1.Enabled = False
    lblTiempo.Visible = False
    lblAutor.Font.Underline = False
    lblAutor.MousePointer = vbDefault
End Sub

Private Sub Timer1_Timer()
    Tiempo = Tiempo - 1
    lblTiempo.Caption = "Tiempo restante: " & Tiempo & "(s)"
    
    If Tiempo <= 0 Then
        Tries = 5
        Timer1.Enabled = False
        Cooldown = False
        cmdInicio.Enabled = True
        lblTiempo.Caption = ""
        MsgBox "Puede iniciar sesión nuevamente", vbInformation, "Cooldown terminado"
    End If
End Sub
