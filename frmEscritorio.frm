VERSION 5.00
Begin VB.Form frmEscritorio 
   BackColor       =   &H80000008&
   Caption         =   "Escritorio"
   ClientHeight    =   10755
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   15990
   BeginProperty Font 
      Name            =   "Noto Sans SC Medium"
      Size            =   8.25
      Charset         =   0
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEscritorio.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmEscritorio.frx":1084A
   MousePointer    =   99  'Custom
   Picture         =   "frmEscritorio.frx":21094
   ScaleHeight     =   10755
   ScaleWidth      =   15990
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   15000
      Top             =   240
   End
   Begin VB.CommandButton cmdApp4 
      Caption         =   "App4"
      Height          =   855
      Left            =   1680
      Picture         =   "frmEscritorio.frx":212CFC
      TabIndex        =   19
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdApp3 
      Caption         =   "App3"
      Height          =   855
      Left            =   600
      Picture         =   "frmEscritorio.frx":223546
      TabIndex        =   18
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdApp2 
      Caption         =   "App2"
      Height          =   855
      Left            =   1680
      TabIndex        =   15
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   240
      Picture         =   "frmEscritorio.frx":233D90
      ScaleHeight     =   675
      ScaleWidth      =   2595
      TabIndex        =   12
      Top             =   3360
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   6240
      Picture         =   "frmEscritorio.frx":23492D
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.Frame frameBotones 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Caption         =   "Botones"
      ForeColor       =   &H00808000&
      Height          =   5295
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   2415
      Begin VB.CommandButton cmdApp1 
         Caption         =   "App1"
         Height          =   855
         Left            =   240
         Picture         =   "frmEscritorio.frx":234E31
         TabIndex        =   14
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox Modo 
         Height          =   360
         Left            =   360
         TabIndex        =   13
         Text            =   "Elige un Modo"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdTerminal 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Terminal"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdMozilla 
         BackColor       =   &H00808000&
         Caption         =   "Mozilla"
         Height          =   615
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdExplorador 
         BackColor       =   &H00C0C000&
         Caption         =   "Explorador Archivos"
         Height          =   615
         Left            =   240
         MouseIcon       =   "frmEscritorio.frx":24567B
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame frameFondo 
      BackColor       =   &H80000004&
      Caption         =   "Fondos de Pantalla"
      ForeColor       =   &H00808000&
      Height          =   3135
      Left            =   5640
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
      Begin VB.OptionButton optFondo4 
         Caption         =   "Fondo No. 4"
         BeginProperty Font 
            Name            =   "Noto Sans SC"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   2400
         Width           =   1335
      End
      Begin VB.OptionButton optFondo3 
         Caption         =   "Fondo No. 3"
         BeginProperty Font 
            Name            =   "Noto Sans SC"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton optFondo2 
         Caption         =   "Fondo No. 2"
         BeginProperty Font 
            Name            =   "Noto Sans SC"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         MousePointer    =   4  'Icon
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optFondo1 
         Caption         =   "Fondo No. 1"
         BeginProperty Font 
            Name            =   "Noto Sans SC"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdBateria 
      Caption         =   "Batería (da CLICK aquí)"
      DisabledPicture =   "frmEscritorio.frx":255EC5
      BeginProperty Font 
         Name            =   "Noto Sans SC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      MouseIcon       =   "frmEscritorio.frx":26670F
      MousePointer    =   99  'Custom
      Picture         =   "frmEscritorio.frx":276F59
      TabIndex        =   16
      Top             =   9240
      Width           =   1455
   End
   Begin VB.Label lblBateria 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   9840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   480
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      FillColor       =   &H00FFFF00&
      Height          =   1215
      Left            =   6000
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      X1              =   120
      X2              =   3840
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   2  'Dash
      Height          =   5535
      Left            =   240
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   1575
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label lblHora 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   855
      Left            =   5520
      TabIndex        =   8
      Top             =   9720
      Width           =   6255
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1695
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   6975
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFC0C0&
      Height          =   975
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   9600
      Width           =   4575
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "Sistema"
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuApps 
      Caption         =   "Aplicaciones"
      Begin VB.Menu mnuBlocDeNotas 
         Caption         =   "Bloc de Notas"
      End
      Begin VB.Menu mnuCalculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu mnuPaint 
         Caption         =   "Paint"
      End
      Begin VB.Menu mnuOffice 
         Caption         =   "Office"
         Begin VB.Menu mnuPowerPoint 
            Caption         =   "PowerPoint"
         End
         Begin VB.Menu mnuExcel 
            Caption         =   "Excel"
         End
         Begin VB.Menu mnuWord 
            Caption         =   "Word"
         End
      End
   End
   Begin VB.Menu mnuPanelControl 
      Caption         =   "Panel de Control"
      Begin VB.Menu mnuFondos 
         Caption         =   "Fondos"
      End
   End
End
Attribute VB_Name = "frmEscritorio"
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

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
    (ByVal lpFileName As String) As Long

Private Function FileExists(ByVal filePath As String) As Boolean
    Dim attr As Long
    attr = GetFileAttributes(filePath)
    
    If attr = -1 Then
        FileExists = False  'si no existe
    Else
        FileExists = (attr And vbDirectory) <> vbDirectory  'si es un directorio
    End If
End Function

Private Sub cmdApp1_Click()
    Dim appPath As String
    Dim user As String
    Dim URL As String
    Dim X As Long
    user = Trim$(Environ$("USERPROFILE"))
    
    Select Case Modo.Text
        Case "Productividad"
            cmdApp1.Caption = "Visual Studio Code"
            
            appPath = user & "\AppData\Local\Programs\Microsoft VS Code\Code.exe"
            URL = "https://code.visualstudio.com/download"
            'Depuracion
            'Label1.Caption = user
            'MsgBox "Ruta: " & appPath
            
            If FileExists(appPath) Then
                'Si existe, abrir app
                X = ShellExecute(0, "open", appPath, vbNullString, vbNullString, 1)
            Else
                'Si no existe, abrir url
                X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            End If
            
            If X <= 32 Then
                MsgBox "Error al ejecutar: " & errorMsg, vbExclamation, "Error"
            End If
            
        Case "Relajación"
            cmdApp1.Caption = "Spotify"
            
            appPath = user & "\AppData\Roaming\Spotify\Spotify.exe"
            URL = "https://www.spotify.com/mx/download/windows/"
            
            If FileExists(appPath) Then
                X = ShellExecute(0, "open", appPath, vbNullString, vbNullString, 1)
            Else
                X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            End If
            
            If X <= 32 Then
                MsgBox "Error al ejecutar: " & errorMsg, vbExclamation, "Error"
            End If
            
        Case "Social"
            cmdApp1.Caption = "WhatsApp"
        
            URL = "https://web.whatsapp.com/"
            X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            If X <= 32 Then
                MsgBox "No se pudo abrir el enlace", vbExclamation, "Error"
            End If
    End Select
End Sub

Private Sub cmdApp2_Click()
    Dim appPath As String
    Dim user As String
    Dim URL As String
    Dim X As Long
    user = Trim$(Environ$("USERPROFILE"))
    
    Select Case Modo.Text
        Case "Productividad"
            cmdApp2.Caption = "Notion"
            
            appPath = user & "\AppData\Local\Programs\Notion\Notion.exe"
            URL = "https://www.notion.com/es/desktop"

            If FileExists(appPath) Then
                X = ShellExecute(0, "open", appPath, vbNullString, vbNullString, 1)
            Else
                X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            End If
            
            If X <= 32 Then
                MsgBox "Error al ejecutar: " & errorMsg, vbExclamation, "Error"
            End If
            
        Case "Relajación"
            cmdApp2.Caption = "Steam"
            
            appPath = "C:\Program Files (x86)\Steam"
            URL = "https://store.steampowered.com/about/?l=spanish"

            If FileExists(appPath) Then
                X = ShellExecute(0, "open", appPath, vbNullString, vbNullString, 1)
            Else
                X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            End If
            
            If X <= 32 Then
                MsgBox "Error al ejecutar: " & errorMsg, vbExclamation, "Error"
            End If
            
        Case "Social"
            cmdApp2.Caption = "Discord"
            
            appPath = user & "\AppData\Local\Discord\Update.exe"
            URL = "https://discord.com/download"

            If FileExists(appPath) Then
                X = ShellExecute(0, "open", appPath, vbNullString, vbNullString, 1)
            Else
                X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            End If
            
            If X <= 32 Then
                MsgBox "Error al ejecutar: " & errorMsg, vbExclamation, "Error"
            End If
    End Select
End Sub

Private Sub cmdApp3_Click()
    Dim appPath As String
    Dim user As String
    Dim URL As String
    Dim X As Long
    user = Trim$(Environ$("USERPROFILE"))
    
    Select Case Modo.Text
        Case "Productividad"
            cmdApp3.Caption = "MongoDB"
            
            appPath = user & "\AppData\Local\MongoDBCompass\MongoDBCompass.exe"
            URL = "https://www.mongodb.com/docs/manual/installation/"
            
            If FileExists(appPath) Then
                X = ShellExecute(0, "open", appPath, vbNullString, vbNullString, 1)
            Else
                X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            End If
            
            If X <= 32 Then
                MsgBox "Error al ejecutar: " & errorMsg, vbExclamation, "Error"
            End If
            
        Case "Relajación"
            cmdApp3.Caption = "Pinterest"
        
            URL = "https://www.google.com/url?sa=t&source=web&rct=j&opi=89978449&url=https://mx.pinterest.com/login/"
            X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            If X <= 32 Then
                MsgBox "No se pudo abrir el enlace", vbExclamation, "Error"
            End If
        Case "Social"
            cmdApp3.Caption = "Telegram"
        
            URL = "https://desktop.telegram.org/"
            X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            If X <= 32 Then
                MsgBox "No se pudo abrir el enlace", vbExclamation, "Error"
            End If
    End Select
End Sub

Private Sub cmdApp4_Click()
    Dim appPath As String
    Dim user As String
    Dim URL As String
    Dim X As Long
    user = Trim$(Environ$("USERPROFILE"))
    
    Select Case Modo.Text
        Case "Productividad"
            cmdApp4.Caption = "Zoom"
            
            appPath = user & "\AppData\Roaming\Zoom\bin\Zoom_launcher.exe"
            URL = "https://zoom.us/es/download"

            If FileExists(appPath) Then
                X = ShellExecute(0, "open", appPath, vbNullString, vbNullString, 1)
            Else
                X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            End If
            
            If X <= 32 Then
                MsgBox "Error al ejecutar: " & errorMsg, vbExclamation, "Error"
            End If
            
        Case "Relajación"
            cmdApp4.Caption = "YouTube"
            
            URL = "https://www.youtube.com/"
            X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            If X <= 32 Then
                MsgBox "No se pudo abrir el enlace", vbExclamation, "Error"
            End If
            
        Case "Social"
            cmdApp4.Caption = "Twitter (X)"
            
            URL = "https://x.com/"
            X = ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
            If X <= 32 Then
                MsgBox "No se pudo abrir el enlace", vbExclamation, "Error"
            End If
    End Select
End Sub

' SE TRABA SI LO PONGO ENTONCES LO QUITÉ
'Private Sub cmdApp1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '   Me.cmdApp1.Enabled = True
  '  Me.cmdApp2.Enabled = False
   ' Me.cmdApp3.Enabled = False
    'Me.cmdApp4.Enabled = False
'End Sub

'Private Sub cmdApp2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '   Me.cmdApp2.Enabled = True
  '  Me.cmdApp1.Enabled = False
   ' Me.cmdApp3.Enabled = False
    'Me.cmdApp4.Enabled = False
'End Sub

'Private Sub cmdApp3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '   Me.cmdApp3.Enabled = True
  '  Me.cmdApp2.Enabled = False
   ' Me.cmdApp1.Enabled = False
    'Me.cmdApp4.Enabled = False
'End Sub

'Private Sub cmdApp4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '   Me.cmdApp4.Enabled = True
  '  Me.cmdApp2.Enabled = False
   ' Me.cmdApp3.Enabled = False
    'Me.cmdApp1.Enabled = False
'End Sub


Private Sub cmdBateria_Click()
    Dim SysInfo As Object
    Set SysInfo = CreateObject("SysInfo.SysInfo")
    
    lblBateria.Caption = SysInfo.BatteryLifePercent & "%"
    lblBateria.Visible = True
End Sub

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
    Me.WindowState = vbMaximized
    lblFecha.FontSize = 30
    lblHora.FontSize = 30
    lblFecha.FontBold = True
    ActualizarTimer
    
    Modo.AddItem "Productividad"
    Modo.AddItem "Relajación"
    Modo.AddItem "Social"
End Sub

Private Sub Form_Resize()
    Me.Refresh
End Sub

Private Sub frameBotones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmdExplorador.Enabled = True
    Me.cmdMozilla.Enabled = True
    Me.cmdTerminal.Enabled = True
    'frameBotones.BackColor = RGB(57, 255, 20)
    'frameBotones.BackColor = vbBlack
End Sub


Private Sub Label1_Click()

End Sub

Private Sub mnuFondos_Click()
    If Me.frameFondo.Visible = False Then
        Me.frameFondo.Visible = True
    Else
        Me.frameFondo.Visible = False
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


Private Sub Modo_Click()
    cmdApp1.Visible = False
    cmdApp2.Visible = False
    cmdApp3.Visible = False
    cmdApp4.Visible = False
    
    Select Case Modo.Text
        Case "Productividad"
            cmdApp1.Caption = "Visual Studio Code"
            cmdApp2.Caption = "Notion"
            cmdApp3.Caption = "MongoDB"
            cmdApp4.Caption = "Zoom"
            
            cmdApp1.Visible = True
            cmdApp2.Visible = True
            cmdApp3.Visible = True
            cmdApp4.Visible = True

        Case "Relajación"
            cmdApp1.Caption = "Spotify"
            cmdApp2.Caption = "Steam"
            cmdApp3.Caption = "Pinterest"
            cmdApp4.Caption = "YouTube"
            
            cmdApp1.Visible = True
            cmdApp2.Visible = True
            cmdApp3.Visible = True
            cmdApp4.Visible = True
        Case "Social"
            cmdApp1.Caption = "WhatsApp"
            cmdApp2.Caption = "Discord"
            cmdApp3.Caption = "Telegram"
            cmdApp4.Caption = "Twitter (X)"
            
            cmdApp1.Visible = True
            cmdApp2.Visible = True
            cmdApp3.Visible = True
            cmdApp4.Visible = True
    End Select
End Sub

Private Sub optFondo1_Click()
    Me.Picture = LoadPicture(App.Path & "\imgs2\alien-movie.jpg")
End Sub

Private Sub optFondo2_Click()
    Me.Picture = LoadPicture(App.Path & "\imgs2\astronaut.jpg")
End Sub

Private Sub optFondo3_Click()
    Me.Picture = LoadPicture(App.Path & "\imgs2\alien-movie2.jpg")
End Sub

Private Sub optFondo4_Click()
    Me.Picture = LoadPicture(App.Path & "\imgs2\wrte.jpg")
End Sub

Private Sub mnuExcel_Click()
    Dim X As Integer
    X = Shell("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")
End Sub

Private Sub mnuWord_Click()
    Dim X As Integer
    X = Shell("C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")
End Sub

Private Sub mnuPowerPoint_Click()
    Dim X As Integer
    X = Shell("C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE")
End Sub

Private Sub mnuCalculadora_Click()
    Dim X As Integer
    X = Shell("C:\Windows\System32\calc.exe")
End Sub

Private Sub mnuPaint_Click()
    Dim respuesta As Integer
    Dim X As Integer
    Dim user As String
    
    user = Trim$(Environ$(userProfile))
    appPath = user & "\AppData\Local\Microsoft\WindowsApps\Microsoft.Paint_8wekyb3d8bbwe\mspaint.exe"
       
    If FileExists(appPath) Then
        X = ShellExecute(0, "open", appPath, vbNullString, vbNullString, 1)
    Else
        MsgBox "Error al ejecutar: " & errorMsg, vbExclamation, "Error"
    End If
End Sub

Private Sub mnuImprimir_Click()
    MsgBox "Error al ejecutar: " & errorMsg, vbExclamation, "Error"
End Sub

Private Sub mnuBlocDeNotas_Click()
    Dim X As Integer
    X = Shell("C:\Windows\System32\notepad.exe", vbNormalFocus)
End Sub

Sub MostrarNotificacion(msg As String)
    MsgBox msg, vbInformation, "Notificación del Sistema"
End Sub

Private Sub ActualizarTimer()
    lblHora.Caption = Format(Now, "hh:mm:ss AM/PM")
    lblFecha.Caption = "Hoy: " & Format(Now, "dddd dd/mm/yyyy")
    Dim SysInfo As Object
    Set SysInfo = CreateObject("SysInfo.SysInfo")
    
    lblBateria.Caption = SysInfo.BatteryLifePercent & "%"
End Sub

Private Sub Picture1_Click()
    MostrarNotificacion "los ALIENS te están buscando !!"
    Me.Picture = LoadPicture(App.Path & "\imgs2\alien-movie2.jpg")
End Sub

Private Sub Timer1_Timer()
    ActualizarTimer
End Sub


