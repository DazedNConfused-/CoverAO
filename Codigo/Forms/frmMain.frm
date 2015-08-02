VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "CoverAO"
   ClientHeight    =   9000
   ClientLeft      =   345
   ClientTop       =   360
   ClientWidth     =   12000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   Picture         =   "frmMain.frx":324A
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   7920
      Top             =   2160
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8640
      TabIndex        =   42
      Top             =   6960
      Width           =   1215
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   10
      Left            =   6090
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   41
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   9
      Left            =   5535
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   40
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   8
      Left            =   4950
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   39
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   7
      Left            =   4350
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   38
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   6
      Left            =   3765
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   37
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   5
      Left            =   3180
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   36
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   2580
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   35
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   1995
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   34
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   1410
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   825
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   0
      Left            =   225
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   31
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox Shermie 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   8040
      ScaleHeight     =   1095
      ScaleWidth      =   735
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
      Begin VB.Label Label5 
         Caption         =   "Msj gm"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Mp"
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Gritar"
         Height          =   255
         Left            =   0
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Global"
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   735
      End
   End
   Begin RichTextLib.RichTextBox RecCombat 
      Height          =   1500
      Left            =   210
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   165
      Visible         =   0   'False
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   99
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":18E17
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RecGlobal 
      Height          =   1500
      Left            =   210
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   165
      Visible         =   0   'False
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   12632319
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":18E94
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      MaxLength       =   500
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1725
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      ItemData        =   "frmMain.frx":18F11
      Left            =   8895
      List            =   "frmMain.frx":18F13
      MousePointer    =   99  'Custom
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   10170
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   7
      Top             =   7335
      Width           =   1500
      Begin VB.Shape UserP 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   45
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   45
      End
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   7440
      Top             =   2160
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   6840
      Top             =   2160
   End
   Begin RichTextLib.RichTextBox RecChat 
      Height          =   1500
      Left            =   210
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   165
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":18F15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   210
      MousePointer    =   99  'Custom
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   544
      TabIndex        =   9
      Top             =   2055
      Width           =   8160
      Begin Captura.wndCaptura Captura1 
         Left            =   240
         Top             =   4920
         _ExtentX        =   688
         _ExtentY        =   688
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   9030
      MousePointer    =   99  'Custom
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   2175
      Width           =   2400
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7800
      TabIndex        =   30
      Top             =   1725
      Width           =   540
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Mapa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   9360
      TabIndex        =   24
      Top             =   7020
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label lblTxtCombat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combate"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   720
      TabIndex        =   23
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label lblTxtDefault 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label lblTxtGlobal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Global"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1560
      TabIndex        =   21
      Top             =   1800
      Width           =   525
   End
   Begin VB.Image cmdMen 
      Height          =   510
      Left            =   10800
      Top             =   1260
      Width           =   1005
   End
   Begin VB.Image CmdLanzar 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "frmMain.frx":18F92
      MousePointer    =   99  'Custom
      Top             =   4920
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   18
      Top             =   7020
      Width           =   3105
   End
   Begin VB.Image modocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":1985C
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image modoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":19C9A
      ToolTipText     =   "Seguro"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image modorol 
      Height          =   255
      Left            =   9645
      Picture         =   "frmMain.frx":1A0D8
      ToolTipText     =   "Modo Rol"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image nomodocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":1A66E
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image nomodoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":1AAAC
      ToolTipText     =   "Seguro"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image nomodorol 
      Height          =   255
      Left            =   9645
      Picture         =   "frmMain.frx":1AEEA
      ToolTipText     =   "Modo Rol"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgClima 
      Height          =   480
      Left            =   6675
      Top             =   8430
      Width           =   1695
   End
   Begin VB.Label lblHAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   17
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblSED 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   16
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblFU 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   15
      Top             =   8340
      Width           =   345
   End
   Begin VB.Label lblAG 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   14
      Top             =   8520
      Width           =   345
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   13
      Top             =   5850
      Width           =   1350
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   12
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblST 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   11
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   8835
      TabIndex        =   6
      Top             =   870
      Width           =   1800
   End
   Begin VB.Label lblNick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NickDelPersonaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   240
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   0
      Left            =   9255
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   2
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   1905
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   495
      Index           =   0
      Left            =   11475
      MousePointer    =   99  'Custom
      Top             =   3405
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10620
      TabIndex        =   4
      Top             =   5745
      Width           =   1110
   End
   Begin VB.Image cmdDropGold 
      Height          =   300
      Left            =   10260
      MousePointer    =   99  'Custom
      Top             =   5670
      Width           =   300
   End
   Begin VB.Image cmdSalir 
      Height          =   255
      Left            =   11580
      Top             =   180
      Width           =   225
   End
   Begin VB.Image cmdMinimizar 
      Height          =   255
      Left            =   11325
      Top             =   180
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   495
      Index           =   1
      Left            =   11475
      MousePointer    =   99  'Custom
      Top             =   2880
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdInfo 
      Height          =   480
      Left            =   10620
      MousePointer    =   99  'Custom
      Top             =   4920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image cmdCon 
      Height          =   510
      Left            =   9690
      Top             =   1260
      Width           =   1110
   End
   Begin VB.Image cmdInv 
      Height          =   510
      Left            =   8580
      Top             =   1245
      Width           =   1110
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10995
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape ExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8835
      Top             =   900
      Width           =   1800
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8745
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape Hpshp 
      BackColor       =   &H00000080&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8745
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   8745
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      Height          =   135
      Left            =   10320
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   135
      Left            =   10320
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Label lblInvInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   615
      Left            =   9000
      TabIndex        =   8
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   1
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Image InvEqu 
      Height          =   4245
      Left            =   8580
      Picture         =   "frmMain.frx":1B328
      Top             =   1245
      Width           =   3240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
Option Explicit
Private Standelf As Boolean
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'*********CONSOLA*********'
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'*********CONSOLA*********'
Private Const VK_SNAPSHOT = &H2C
'*********FOTO***********'

Public InMouseExp As Boolean

Private LoadC As Boolean
Private LastI As Byte
Private SelectI As Byte

Public tX As Byte
Public tY As Byte
Public MouseX As Long  'SI no pones mod CoverAO te re kbe el juicio Conchatumadre - xD ?)
Public MouseY As Long
Public MouseBoton As Long

Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private Sub cmdCon_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadInterface("Hechizos")

    picInv.Visible = False

    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
    
    cmdMoverHechi(0).Enabled = True
    cmdMoverHechi(1).Enabled = True
    
    Image1(0).Visible = False
    Image1(1).Visible = False
    Image1(2).Visible = False
    
    lblInvInfo.Visible = False
End Sub


Private Sub cmdINV_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadInterface("Inventory")
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
    cmdMoverHechi(0).Enabled = False
    cmdMoverHechi(1).Enabled = False
    lblInvInfo.Visible = True
    
    Image1(0).Visible = False
    Image1(1).Visible = False
    Image1(2).Visible = False
    
    RenderInv = True
End Sub


Private Sub cmdMen_Click()
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadInterface("Menu")
    picInv.Visible = False

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
    
    cmdMoverHechi(0).Enabled = False
    cmdMoverHechi(1).Enabled = False
    
    lblInvInfo.Visible = False
    
    Image1(0).Visible = True
    Image1(1).Visible = True
    Image1(2).Visible = True
End Sub

Private Sub cmdMinimizar_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)
    
    If hlst.ListIndex = -1 Then Exit Sub
    Dim sTemp As String

    Select Case Index
        Case 1 'subir
            If hlst.ListIndex = 0 Then Exit Sub
        Case 0 'bajar
            If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
    End Select

    Call WriteMoveSpell(Index, hlst.ListIndex + 1)
    
    Select Case Index
        Case 1 'subir
            sTemp = hlst.List(hlst.ListIndex - 1)
            hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex - 1
        Case 0 'bajar
            sTemp = hlst.List(hlst.ListIndex + 1)
            hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
            hlst.List(hlst.ListIndex) = sTemp
            hlst.ListIndex = hlst.ListIndex + 1
    End Select
End Sub


Public Sub ControlSeguroResu(ByVal Mostrar As Boolean)
If Mostrar Then
    'If Not PicResu.Visible Then
    '    PicResu.Visible = True
    'End If
Else
    'If PicResu.Visible Then
    '    PicResu.Visible = False
    'End If
End If
End Sub


Public Sub DibujarSeguro()
modoseguro.Visible = True
nomodoseguro.Visible = False
End Sub

Public Sub DesDibujarSeguro()
modoseguro.Visible = False
nomodoseguro.Visible = True
End Sub


Private Sub cmdSalir_Click()
End
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If (Not SendTxt.Visible) Then
        
        'Checks if the key is valid
            Select Case KeyCode
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    Audio.MusicActivated = Not Audio.MusicActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
                    Call WriteCombatModeToggle
                    IScombate = Not IScombate
                    If IScombate = True Then
                        modocombate.Visible = True
                        nomodocombate.Visible = False
                    Else
                        modocombate.Visible = False
                        nomodocombate.Visible = True
                    End If
                    
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call WriteWork(eSkill.Domar)
                
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    Call WriteWork(eSkill.Robar)
                            
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    Call WriteWork(eSkill.Ocultarse)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)

                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
            End Select
        End If
    

    Select Case KeyCode
        Case vbKeyF1 To vbKeyF11
           Call frmBindKey.Bind_Accion(KeyCode - vbKeyF1 + 1)
            
       Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
        Dim i As Integer
        Captura1.Area = Ventana
        Captura1.Captura
        For i = 1 To 1000
            If Not FileExist(App.Path & "\screenshots\Imagen" & i & ".bmp", vbNormal) Then Exit For
        Next
        Call SavePicture(Captura1.Imagen, App.Path & "/screenshots/Imagen" & i & ".bmp")
        Call AddtoRichTextBox(frmMain.RecChat, "Screenshot grabada correctamente como " & i & ".bmp. Subi tu foto ahora a www.imagehack.us", 0, 191, 128, False, False, False)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
            FPSFLAG = Not FPSFLAG

        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            Call WriteMeditate
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If macrotrabajo.Enabled Then
                DesactivarMacroTrabajo
            Else
                ActivarMacroTrabajo
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            If frmMain.macrotrabajo.Enabled Then DesactivarMacroTrabajo
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub

            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
            If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            If macrotrabajo.Enabled Then DesactivarMacroTrabajo
            
            Call WriteAttack
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
              (Not frmBancoObj.Visible) And _
              (Not frmMSG.Visible) And (Not frmForo.Visible) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If

    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    InMouseExp = False
    If UserPasarNivel = 0 Then
        lblExp.Caption = "¡Nivel máximo!"
    Else
        frmMain.lblExp.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%"
    End If
End Sub

Private Sub Label1_Click()
frmMain.SendTxt.SelText = ";"
frmMain.Shermie.Visible = False
End Sub

Private Sub Label3_Click()
frmMain.SendTxt.SelText = "-"
frmMain.Shermie.Visible = False
End Sub

Private Sub Label4_Click()
frmMain.SendTxt.SelText = "\nick"
frmMain.Shermie.Visible = False
End Sub

Private Sub Label5_Click()
frmMain.SendTxt.SelText = "/denunciar"
frmMain.Shermie.Visible = False
End Sub

Private Sub Label6_Click()
frmMain.Shermie.Visible = True
End Sub

Private Sub lblExp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
InMouseExp = True
lblExp.Caption = UserExp & "/" & UserPasarNivel
If UserPasarNivel = 0 Then
    lblExp.Caption = "¡Nivel máximo!"
End If
End Sub

Private Sub lblTxtCombat_Click()
    RecChat.Visible = False
    RecCombat.Visible = True
    RecGlobal.Visible = False
    
    lblTxtGlobal.font.bold = False
    lblTxtDefault.font.bold = False
    lblTxtCombat.font.bold = True
End Sub
Private Sub lblTxtDefault_Click()
    RecChat.Visible = True
    RecCombat.Visible = False
    RecGlobal.Visible = False
    
    lblTxtGlobal.font.bold = False
    lblTxtDefault.font.bold = True
    lblTxtCombat.font.bold = False
End Sub
Private Sub lblTxtGlobal_Click()
    RecChat.Visible = False
    RecCombat.Visible = False
    RecGlobal.Visible = True
    
    lblTxtGlobal.font.bold = True
    lblTxtDefault.font.bold = False
    lblTxtCombat.font.bold = False
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    clicX = x
    clicY = y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub


Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    If Not IsAppActive() Then
        DesactivarMacroTrabajo
        Exit Sub
    End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.Visible) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     If Not (frmCarp.Visible = True) Then Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecChat, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, False)
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0

    Call AddtoRichTextBox(frmMain.RecChat, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, False)
End Sub


Private Sub Minimap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call WriteWarpChar("YO", UserMap, IIf(x < 1, 1, x), y)
    DibujarMiniMapPos
End Sub


Private Sub picMacro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
BotonElegido = Index + 1

If MacroKeys(BotonElegido).TipoAccion = 0 Or Button = vbRightButton Then
    frmBindKey.Show vbModeless, frmMain
Else
    Call frmBindKey.Bind_Accion(Index + 1)
End If
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub

Private Sub Second_Timer()
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
    With luz_dia(Hour(time))
        base_light = Engine.change_day_effect(day_r_old, day_g_old, day_b_old, .r, .g, .b)
    End With
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            Call WriteDrop(Inventario.SelectedItem, 1)
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    Call WritePickUp
End Sub

Private Sub UsarItem()

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''


Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(Vacio)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub



Private Sub MainViewPic_Click()

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                
                    If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        RestaurarIcon
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            'Call AddtoRichTextBox(frmMain.RecChat, "No podés lanzar flechas tan rapido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            RestaurarIcon
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                'Call AddtoRichTextBox(frmMain.RecChat, "No podés lanzar flechas tan rapido.", .red, .green, .blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                RestaurarIcon
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                   ' Call AddtoRichTextBox(frmMain.RecChat, "No puedes lanzar hechizos tan rápido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                RestaurarIcon
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    'Call AddtoRichTextBox(frmMain.RecChat, "No podés lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            RestaurarIcon
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If

                    RestaurarIcon
                    
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                If Not frmForo.Visible And Not frmComerciar.Visible And Not frmBancoObj.Visible Then
                    Call WriteDoubleClick(tX, tY)
                End If
            End If
        ElseIf MouseBoton = vbRightButton Then
            Call WriteDoubleClick(tX, tY)
        ElseIf (MouseShift And 1) = 1 Then
            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
    End If
End Sub

Private Sub MainViewPic_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
    If Not frmForo.Visible And Not frmComerciar.Visible And Not frmBancoObj.Visible Then
     '  Call WriteDoubleClick(tX, tY)
    End If
End Sub

Private Sub Form_Load()

    SetWindowLong RecChat.hwnd, -20, &H20& 'Consola Transparente
    SetWindowLong RecGlobal.hwnd, -20, &H20&
    SetWindowLong RecCombat.hwnd, -20, &H20&
    
    frmMain.Caption = "CoverAO"
    Detectar RecChat.hwnd, Me.hwnd
    
    Me.Picture = LoadInterface("Main")
    InvEqu.Picture = LoadInterface("Inventory")
    
    Me.Left = 0
    Me.Top = 0
   
    Dim CursorDir As String
    Dim Cursor As Long
    
    CursorDir = App.Path & "\Recursos\Main.cur" 'Shermie80/maycolito (:
    hSwapCursor = SetClassLong(frmMain.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    hSwapCursor = SetClassLong(frmMain.MainViewPic.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    hSwapCursor = SetClassLong(frmMain.hlst.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorDir))
    
    lblTxtGlobal.font.bold = False
    lblTxtDefault.font.bold = True
    lblTxtCombat.font.bold = False
    RecChat.Visible = True
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x
    MouseY = y
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
        Case 0
            Call frmOpciones.Show(vbModeless, frmMain)
            
        Case 1
            LlegaronEstadisticas = False
            Call WriteRequestEstadisticas
            Call FlushBuffer
            
            Do While Not LlegaronEstadisticas
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronEstadisticas = False
        
        Case 2
            If frmGuildLeader.Visible Then Unload frmGuildLeader
            
            Call WriteRequestGuildLeaderInfo
    End Select
End Sub

Private Sub cmdDropGold_Click()
    Inventario.SelectGold
    If UserGLD > 0 Then
        frmCantidad.Show , frmMain
    End If
End Sub


Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
     If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otMapas Then
            frmMap.Show , frmMain
            frmMap.Top = frmMain.Top
            frmMap.Left = frmMain.Left
     End If
    
    If macrotrabajo.Enabled Then DesactivarMacroTrabajo
    Call EquiparItem
    Call UsarItem

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (Not frmComerciar.Visible) And (Not frmComerciarUsu.Visible) And _
      (Not frmBancoObj.Visible) And _
      (Not frmMSG.Visible) And (Not frmForo.Visible) And _
      (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) And (picInv.Visible) Then
        picInv.SetFocus
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub



''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
Private Sub Socket1_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)
    
    Second.Enabled = True

    Call Login
    
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    Dim mifrm As Form
    
    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    On Local Error Resume Next
    For Each mifrm In Forms
        If Not mifrm.name = Me.name And Not mifrm.name = frmCrearPersonaje.name And Not mifrm.name = frmConnect.name Then
            Unload mifrm
        End If
    Next
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False
    
    Pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    macrotrabajo.Enabled = False

    SkillPoints = 0
    Alocados = 0
    
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Si el servidor no le conecta en unos minutos, tiene problemas con su internet, por favor verifìquelo", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    ElseIf ErrorCode = 24061 Then
        Call MsgBox("No hay coneccion con el servidor. Porfavor verifique su estado o bien su coneccion de internet.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")

    Response = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect

End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    Debug.Print Asc(RD)
    data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub
    
    
    'Put data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub
Sub RestaurarIcon()
    Me.MousePointer = 99
End Sub
