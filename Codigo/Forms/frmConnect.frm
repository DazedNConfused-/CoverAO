VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmConnect 
   BorderStyle     =   0  'None
   Caption         =   "Cover AO"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmConnect.frx":324A
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Passtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Usertxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   390
      Left            =   2160
      TabIndex        =   0
      Top             =   2040
      Width           =   4335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3135
      Left            =   2160
      TabIndex        =   2
      Top             =   5160
      Width           =   7635
      ExtentX         =   13467
      ExtentY         =   5530
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ListBox lst_servers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      ItemData        =   "frmConnect.frx":53DB1
      Left            =   6600
      List            =   "frmConnect.frx":53DB3
      TabIndex        =   1
      Top             =   2040
      Width           =   3165
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   11400
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11640
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Image cmdConnect 
      Height          =   630
      Left            =   4770
      Top             =   2700
      Width           =   1800
   End
   Begin VB.Image cmdNewAccount 
      Height          =   660
      Left            =   2250
      Top             =   3720
      Width           =   2100
   End
   Begin VB.Image cmdNotReme 
      Height          =   660
      Left            =   4455
      Top             =   3720
      Width           =   2100
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdConnect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdConnect.Picture = LoadInterface("btnConectarApretado")
End Sub
Private Sub cmdConnect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 0 Then cmdConnect.Picture = LoadInterface("btnConectarMouse")
    Set cmdNewAccount.Picture = Nothing
    Set cmdNotReme.Picture = Nothing
End Sub

Private Sub cmdNewAccount_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdNewAccount.Picture = LoadInterface("btnCuentaApretado")
End Sub

Private Sub cmdNewAccount_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 0 Then cmdNewAccount.Picture = LoadInterface("btnCuentaMouse")
    Set cmdConnect.Picture = Nothing
    Set cmdNotReme.Picture = Nothing
End Sub

Private Sub cmdNotReme_Click()
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub cmdNotReme_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdNotReme.Picture = LoadInterface("btnPassApretado")

End Sub

Private Sub cmdNotReme_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 0 Then cmdNotReme.Picture = LoadInterface("btnPassMouse")
    Set cmdConnect.Picture = Nothing
    Set cmdNewAccount.Picture = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        bRunning = False
    End If
End Sub

Private Sub Form_Load()
    Me.Picture = LoadInterface("Conectar")
    Me.Icon = frmMain.Icon
    
    lServer(1).port = 7666
    lServer(1).Ip = "127.0.0.1"
    lServer(1).name = "Larias(Exp x25 Orox20)(Offine)"
    
    lServer(2).port = 7666
    lServer(2).Ip = "127.0.0.1"
    lServer(2).name = "Alphaeron(Exp x50 Oro x40)(online)"
    
    lst_servers.AddItem lServer(1).name
    lst_servers.AddItem lServer(2).name

    WebBrowser1.Navigate "http://coverAO.jimdo.com"

    lst_servers.ListIndex = 1

    Usertxt.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set cmdConnect.Picture = Nothing
    Set cmdNewAccount.Picture = Nothing
    Set cmdNotReme.Picture = Nothing
End Sub

Private Sub cmdNewAccount_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call Audio.PlayMIDI("7.mid")
    frmCrearCuenta.Show
End Sub

Private Sub cmdConnect_Click()
Call Audio.PlayWave(SND_CLICK)
    
If frmMain.Socket1.Connected Then
    frmMain.Socket1.Disconnect
    frmMain.Socket1.Cleanup
    DoEvents
End If
    
UserAccount = Usertxt.Text
UserPassword = passTxt.Text

If Not UserAccount = "" And Not UserPassword = "" Then
    EstadoLogin = ConectarCuenta
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
End If

End Sub

Private Sub Label1_Click()
End
End Sub

Private Sub Label2_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub lst_servers_Click()
    CurServerIp = lServer(lst_servers.ListIndex + 1).Ip
    CurServerPort = lServer(lst_servers.ListIndex + 1).port
End Sub

