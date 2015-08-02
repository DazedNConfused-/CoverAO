VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   Icon            =   "frmCrearPersonaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonaje.frx":0CCA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox lstFamiliar 
      BackColor       =   &H00000000&
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
      Height          =   285
      ItemData        =   "frmCrearPersonaje.frx":646FB
      Left            =   8520
      List            =   "frmCrearPersonaje.frx":646FD
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   1845
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox txtFamiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9345
      MaxLength       =   20
      TabIndex        =   46
      Top             =   975
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox picFamiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   10380
      ScaleHeight     =   1185
      ScaleWidth      =   840
      TabIndex        =   45
      Top             =   1575
      Width           =   870
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":646FF
      Left            =   870
      List            =   "frmCrearPersonaje.frx":64701
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   3810
      Width           =   2055
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":64703
      Left            =   870
      List            =   "frmCrearPersonaje.frx":64705
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   3150
      Width           =   2055
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":64707
      Left            =   870
      List            =   "frmCrearPersonaje.frx":64709
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   2490
      Width           =   2055
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
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
      Height          =   495
      Left            =   2100
      MaxLength       =   30
      TabIndex        =   41
      Top             =   1050
      Width           =   5865
   End
   Begin VB.PictureBox HeadView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1695
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   40
      Top             =   4545
      Width           =   375
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":6470B
      Left            =   8520
      List            =   "frmCrearPersonaje.frx":64712
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3600
      Width           =   2820
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   525
      Left            =   2640
      TabIndex        =   49
      Top             =   8280
      Width           =   6795
   End
   Begin VB.Image imgNoDisp 
      Height          =   2145
      Left            =   8415
      Top             =   795
      Width           =   3045
   End
   Begin VB.Label lblFamiInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descropcion del familiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   8535
      TabIndex        =   48
      Top             =   2235
      Width           =   1635
   End
   Begin VB.Image boton 
      Height          =   615
      Index           =   0
      Left            =   9600
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   8160
      Width           =   1755
   End
   Begin VB.Image boton 
      Height          =   615
      Index           =   1
      Left            =   645
      MousePointer    =   99  'Custom
      Tag             =   "0"
      Top             =   8160
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   3570
      Left            =   8475
      Stretch         =   -1  'True
      Top             =   4260
      Width           =   2835
   End
   Begin VB.Image MenosHead 
      Height          =   600
      Left            =   1320
      Tag             =   "0"
      Top             =   4440
      Width           =   390
   End
   Begin VB.Image MasHead 
      Height          =   600
      Left            =   2160
      Tag             =   "0"
      Top             =   4440
      Width           =   390
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2460
      TabIndex        =   39
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label lbAtributos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2535
      TabIndex        =   38
      Top             =   7500
      Width           =   255
   End
   Begin VB.Label modfuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   37
      Top             =   5700
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   0
      Left            =   2715
      Top             =   5640
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   0
      Left            =   2715
      Top             =   5790
      Width           =   195
   End
   Begin VB.Label modAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   36
      Top             =   6060
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label modInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   35
      Top             =   6420
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label modCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   34
      Top             =   6780
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label modConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   33
      Top             =   7140
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2460
      TabIndex        =   32
      Top             =   6060
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2460
      TabIndex        =   31
      Top             =   6420
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2460
      TabIndex        =   30
      Top             =   6780
      Width           =   240
   End
   Begin VB.Label lbAtt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2460
      TabIndex        =   29
      Top             =   7140
      Width           =   240
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   1
      Left            =   2715
      Top             =   6030
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   2
      Left            =   2715
      Top             =   6360
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   3
      Left            =   2715
      Top             =   6750
      Width           =   195
   End
   Begin VB.Image ImgAtributoMas 
      Height          =   135
      Index           =   4
      Left            =   2715
      Top             =   7080
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   1
      Left            =   2715
      Top             =   6150
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   2
      Left            =   2715
      Top             =   6510
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   3
      Left            =   2715
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image ImgAtributoMenos 
      Height          =   135
      Index           =   4
      Left            =   2715
      Top             =   7230
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   1
      Left            =   5310
      TabIndex        =   28
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   0
      Left            =   5310
      TabIndex        =   27
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   2
      Left            =   5310
      TabIndex        =   26
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   3
      Left            =   5310
      TabIndex        =   25
      Top             =   3450
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   4
      Left            =   5310
      TabIndex        =   24
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   5
      Left            =   5310
      TabIndex        =   23
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   6
      Left            =   5310
      TabIndex        =   22
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   7
      Left            =   5310
      TabIndex        =   21
      Top             =   4950
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   8
      Left            =   5310
      TabIndex        =   20
      Top             =   5340
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   9
      Left            =   5310
      TabIndex        =   19
      Top             =   5700
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   10
      Left            =   5310
      TabIndex        =   18
      Top             =   6090
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   11
      Left            =   5310
      TabIndex        =   17
      Top             =   6450
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   12
      Left            =   5310
      TabIndex        =   16
      Top             =   6840
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   13
      Left            =   5310
      TabIndex        =   15
      Top             =   7200
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   14
      Left            =   7365
      TabIndex        =   14
      Top             =   2340
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   15
      Left            =   7365
      TabIndex        =   13
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   16
      Left            =   7365
      TabIndex        =   12
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   17
      Left            =   7365
      TabIndex        =   11
      Top             =   3450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":6471B
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":6486D
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":649BF
      Top             =   4320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":64B11
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":64C63
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":64DB5
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":64F07
      Top             =   3570
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":65059
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":651AB
      Top             =   2820
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":652FD
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":6544F
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":655A1
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":656F3
      Top             =   7170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":65845
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":65997
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":65AE9
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":65C3B
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":65D8D
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":65EDF
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":66031
      Top             =   4560
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":66183
      Top             =   4170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":662D5
      Top             =   3780
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":66427
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":66579
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   2
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":666CB
      Top             =   2670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":6681D
      Top             =   2280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":6696F
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":66AC1
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":66C13
      Top             =   6930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":66D65
      Top             =   6540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":66EB7
      Top             =   6180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":67009
      Top             =   5790
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":6715B
      Top             =   5430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":672AD
      Top             =   5040
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":673FF
      Top             =   4680
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":67551
      Top             =   4290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":676A3
      Top             =   3930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":677F5
      Top             =   3540
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   5
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":67947
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   5565
      MouseIcon       =   "frmCrearPersonaje.frx":67A99
      Top             =   2790
      Width           =   195
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6795
      TabIndex        =   10
      Top             =   7260
      Width           =   255
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   20
      Left            =   7365
      TabIndex        =   9
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   19
      Left            =   7365
      TabIndex        =   8
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   18
      Left            =   7365
      TabIndex        =   7
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   21
      Left            =   7365
      TabIndex        =   6
      Top             =   4950
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":67BEB
      Top             =   4920
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":67D3D
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   22
      Left            =   7365
      TabIndex        =   5
      Top             =   5340
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   44
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":67E8F
      Top             =   5280
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   45
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":67FE1
      Top             =   5430
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   23
      Left            =   7365
      TabIndex        =   4
      Top             =   5700
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   46
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":68133
      Top             =   5670
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   47
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":68285
      Top             =   5790
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   24
      Left            =   7365
      TabIndex        =   3
      Top             =   6090
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   48
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":683D7
      Top             =   6060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   49
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":68529
      Top             =   6180
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   25
      Left            =   7365
      TabIndex        =   2
      Top             =   6450
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   50
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":6867B
      Top             =   6420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   51
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":687CD
      Top             =   6540
      Width           =   195
   End
   Begin VB.Label Skill 
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
      Height          =   195
      Index           =   26
      Left            =   7365
      TabIndex        =   1
      Top             =   6840
      Width           =   240
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   52
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":6891F
      Top             =   6810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   53
      Left            =   7620
      MouseIcon       =   "frmCrearPersonaje.frx":68A71
      Top             =   6930
      Width           =   195
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SkillPoints As Byte
Public Actual As Integer
Private MaxEleccion As Integer, MinEleccion As Integer

Function CheckData() As Boolean
If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = 0 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = 0 Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = 0 Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

'If frmCrearPersonaje.lstFamiliar.Visible = True Then
'    If UserPet.Tipo = "" Then
'        lblInfo.Caption = "Seleccione su familiar o mascota."
'        Exit Function'
'    ElseIf UserPet.nombre = "" Then
'        lblInfo.Caption = "Asigne un nombre a su familiar o mascota."
'        Exit Function
'    ElseIf Len(UserPet.nombre) > 30 Then
'        lblInfo.Caption = ("El nombre de tu familiar o mascota debe tener menos de 30 letras.")
'        Exit Function
'    End If
'End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True

End Function

Private Sub boton_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)
    
    Select Case Index
        Case 0
            
            Dim i As Integer
            Dim k As Object
            i = 1
            For Each k In Skill
                UserSkills(i) = k.Caption
                i = i + 1
            Next
            
            UserName = txtNombre.Text
            
            If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
            End If
            
            UserRaza = lstRaza.ListIndex + 1
            UserSexo = lstGenero.ListIndex + 1
            UserClase = lstProfesion.ListIndex + 1
            
            UserAtributos(1) = Val(lbAtt(0).Caption)
            UserAtributos(2) = Val(lbAtt(1).Caption)
            UserAtributos(3) = Val(lbAtt(2).Caption)
            UserAtributos(4) = Val(lbAtt(3).Caption)
            UserAtributos(5) = Val(lbAtt(4).Caption)
            
            UserHogar = lstHogar.ListIndex + 1

            If CheckData() Then
                frmMain.Socket1.HostName = CurServerIp
                frmMain.Socket1.RemotePort = CurServerPort
                EstadoLogin = CrearNuevoPj
                    
                If Not frmMain.Socket1.Connected Then
                    MsgBox "Error: Se ha perdido la conexion con el server."
                    Unload Me
                Else
                    Call Login
                End If
            End If
            
        Case 1
            Call Audio.PlayMIDI("2.mid")
            Unload Me

    End Select
End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub Command1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If (Index And &H1) = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
Unload frmConnect
SkillPoints = 10
Puntos.Caption = SkillPoints
Me.Picture = LoadInterface("CreaPJ")
'imgHogar.Picture = LoadInterface("CP-Ullathorpe")
Me.Icon = frmMain.Icon

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstHogar.Clear

For i = LBound(Ciudades()) To UBound(Ciudades())
    lstHogar.AddItem Ciudades(i)
Next i


lstRaza.Clear

For i = LBound(ListaRazas()) To UBound(ListaRazas())
    lstRaza.AddItem ListaRazas(i)
Next i


lstProfesion.Clear

For i = LBound(ListaClases()) To UBound(ListaClases())
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.ListIndex = 1

lstGenero.AddItem "Hombre"
lstGenero.AddItem "Mujer"
Image1.Picture = LoadInterface(lstProfesion.Text & "")



End Sub
Private Sub ImgAtributoMas_Click(Index As Integer)

If Val(lbAtt(Index).Caption) >= 18 Or Val(lbAtributos.Caption) <= 0 Then Exit Sub
    
lbAtt(Index).Caption = Val(lbAtt(Index).Caption) + 1
lbAtributos.Caption = lbAtributos.Caption - 1

End Sub

Private Sub ImgAtributoMenos_Click(Index As Integer)

If Val(lbAtt(Index).Caption) <= 6 Then Exit Sub

lbAtt(Index).Caption = Val(lbAtt(Index).Caption) - 1
lbAtributos.Caption = lbAtributos.Caption + 1

End Sub

Private Sub lstFamiliar_Click()
If lstFamiliar.ListIndex > 0 Then
    lblFamiInfo.Caption = ListaFamiliares(lstFamiliar.ListIndex).Desc
    picFamiliar.Picture = LoadInterface(ListaFamiliares(lstFamiliar.ListIndex).Imagen)
Else
    lblFamiInfo.Caption = "Selecciona tu familiar o mascota para saber más de él"
    picFamiliar.Picture = Nothing
End If
End Sub


Private Sub lstProfesion_Click()
On Error Resume Next

Image1.Picture = LoadInterface("" & lstProfesion.Text & "")

If lstProfesion.Text = "Mago" Then
    frmCrearPersonaje.txtFamiliar.Visible = True
    frmCrearPersonaje.lstFamiliar.Visible = True
    imgNoDisp.Picture = Nothing
    lblFamiInfo.Visible = True
    picFamiliar.Visible = True
    Call CambioFamiliar(5)
ElseIf lstProfesion.Text = "Cazador" Or lstProfesion.Text = "Druida" Then
    frmCrearPersonaje.txtFamiliar.Visible = True
    frmCrearPersonaje.lstFamiliar.Visible = True
    imgNoDisp.Picture = Nothing
    lblFamiInfo.Visible = True
    picFamiliar.Visible = True
    Call CambioFamiliar(4)
Else
    frmCrearPersonaje.txtFamiliar.Visible = False
    frmCrearPersonaje.lstFamiliar.Visible = False
    imgNoDisp.Picture = LoadInterface("mascotanodisp.bmp")
    picFamiliar.Visible = False
    lblFamiInfo.Visible = False
End If

End Sub
Private Sub lstGenero_Click()
    Call DameOpciones
End Sub
Private Sub lstRaza_Click()
    Call DameOpciones
        modfuerza.Visible = True
        modConstitucion.Visible = True
        modAgilidad.Visible = True
        modInteligencia.Visible = True
        modCarisma.Visible = True
    Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+1"
        modConstitucion.Caption = "+2"
        modAgilidad.Caption = "+1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+1"
        modAgilidad.Caption = "+3"
        modInteligencia.Caption = "+1"
        modCarisma.Caption = "+2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+1"
        modInteligencia.Caption = "+2"
        modCarisma.Caption = "-3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "-6"
        modCarisma.Caption = "-3"
    Case Is = "Gnomo"
        modfuerza.Caption = "-5"
        modAgilidad.Caption = "+4"
        modInteligencia.Caption = "+3"
        modCarisma.Caption = "+1"
    Case Is = "Orco"
        modfuerza.Caption = "+ 5"
        modConstitucion.Caption = "+6"
        modAgilidad.Caption = "- 2"
        modInteligencia.Caption = "-6"
        modCarisma.Caption = "-2"
End Select
End Sub
Private Sub MenosHead_Click()
Call Audio.PlayWave(SND_CLICK)
Actual = Actual - 1
If Actual > MaxEleccion Then
   Actual = MaxEleccion
ElseIf Actual < MinEleccion Then
   Actual = MinEleccion
End If
HeadView.Cls
Call Engine.DrawGrhToHdc(HeadView.hdc, HeadData(Actual).Head(3).grhindex, 8, 5)
HeadView.Refresh
End Sub
Private Sub MasHead_Click()
Call Audio.PlayWave(SND_CLICK)
Actual = Actual + 1
If Actual > MaxEleccion Then
   Actual = MaxEleccion
ElseIf Actual < MinEleccion Then
   Actual = MinEleccion
End If
HeadView.Cls
Call Engine.DrawGrhToHdc(HeadView.hdc, HeadData(Actual).Head(3).grhindex, 5, 5)
HeadView.Refresh
End Sub

Private Sub CambioFamiliar(ByVal NumFamiliares As Integer)
If NumFamiliares = 5 Then
    ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
    ListaFamiliares(1).name = "Elemental De Fuego"
    ListaFamiliares(1).Desc = "Hecho de puro fuego, lanzará tormentas sobre tus contrincantes."
    ListaFamiliares(1).Imagen = "elefuego"
    
    ListaFamiliares(2).name = "Elemental De Agua"
    ListaFamiliares(2).Desc = "Con su cuerpo acuoso paralizará a tus enemigos."
    ListaFamiliares(2).Imagen = "eleagua"
    
    ListaFamiliares(3).name = "Elemental De Tierra"
    ListaFamiliares(3).Desc = "Sus fuertes brazos inmovilizarán cualquier criatura viviente."
    ListaFamiliares(3).Imagen = "eletierra"
    
    ListaFamiliares(4).name = "Ely"
    ListaFamiliares(4).Desc = "Te protegerá constantemente con sus conjuros defensivos."
    ListaFamiliares(4).Imagen = "ely"
    
    ListaFamiliares(5).name = "Fuego Fatuo"
    ListaFamiliares(5).Desc = "Débil pero con gran poder mágico, siempre estará a tu lado."
    ListaFamiliares(5).Imagen = "fatuo"
Else
    ReDim ListaFamiliares(1 To NumFamiliares) As tListaFamiliares
    ListaFamiliares(1).name = "Tigre"
    ListaFamiliares(1).Desc = "Poseen grandes y filosas garras para atacar a tus oponentes."
    ListaFamiliares(1).Imagen = "tigre"
    
    ListaFamiliares(2).name = "Lobo"
    ListaFamiliares(2).Desc = "Astutos y arrogantes, su mordedura causa estragos en sus víctimas."
    ListaFamiliares(2).Imagen = "lobo"
    
    ListaFamiliares(3).name = "Oso Pardo"
    ListaFamiliares(3).Desc = "Se caracterizan por ser territoriales y muy resistentes."
    ListaFamiliares(3).Imagen = "oso"
    
    ListaFamiliares(4).name = "Ent"
    ListaFamiliares(4).Desc = "¡Esta robusta criatura te defenderá cual muro de piedra!"
    ListaFamiliares(4).Imagen = "ent"
End If
Dim i As Integer
lstFamiliar.Clear
lstFamiliar.AddItem ""
For i = 1 To UBound(ListaFamiliares)
    lstFamiliar.AddItem ListaFamiliares(i).name
Next i
lstFamiliar.ListIndex = 0
End Sub

Sub DameOpciones()
 
Dim i As Integer
 
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = Head_Range(HUMANO).mStart
                MaxEleccion = Head_Range(HUMANO).mEnd
                MinEleccion = Head_Range(HUMANO).mStart
            Case "Elfo"
                Actual = Head_Range(ELFO).mStart
                MaxEleccion = Head_Range(ELFO).mEnd
                MinEleccion = Head_Range(ELFO).mStart
            Case "Drow"
                Actual = Head_Range(ElfoOscuro).mStart
                MaxEleccion = Head_Range(ElfoOscuro).mEnd
                MinEleccion = Head_Range(ElfoOscuro).mStart
            Case "Enano"
                Actual = Head_Range(Enano).mStart
                MaxEleccion = Head_Range(Enano).mEnd
                MinEleccion = Head_Range(Enano).mStart
            Case "Gnomo"
                Actual = Head_Range(Gnomo).mStart
                MaxEleccion = Head_Range(Gnomo).mEnd
                MinEleccion = Head_Range(Gnomo).mStart
            Case "Orco"
                Actual = Head_Range(Orco).mStart
                MaxEleccion = Head_Range(Orco).mEnd
                MinEleccion = Head_Range(Orco).mStart
            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = Head_Range(HUMANO).fStart
                MaxEleccion = Head_Range(HUMANO).fEnd
                MinEleccion = Head_Range(HUMANO).fStart
            Case "Elfo"
                Actual = Head_Range(ELFO).fStart
                MaxEleccion = Head_Range(ELFO).fEnd
                MinEleccion = Head_Range(ELFO).fStart
            Case "Drow"
                Actual = Head_Range(ElfoOscuro).fStart
                MaxEleccion = Head_Range(ElfoOscuro).fEnd
                MinEleccion = Head_Range(ElfoOscuro).fStart
            Case "Enano"
                Actual = Head_Range(Enano).fStart
                MaxEleccion = Head_Range(Enano).fEnd
                MinEleccion = Head_Range(Enano).fStart
            Case "Gnomo"
                Actual = Head_Range(Gnomo).fStart
                MaxEleccion = Head_Range(Gnomo).fEnd
                MinEleccion = Head_Range(Gnomo).fStart
            Case "Orco"
                Actual = Head_Range(Orco).fStart
                MaxEleccion = Head_Range(Orco).fEnd
                MinEleccion = Head_Range(Orco).fStart
            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
End Select
 
HeadView.Cls
Call Engine.DrawGrhToHdc(HeadView.hdc, HeadData(Actual).Head(3).grhindex, 5, 5)
HeadView.Refresh
End Sub
Public Function BonificadorRaza(ByVal Atributo As Integer, ByVal Raza As Byte) As Integer

Select Case Atributo
    Case Fuerza
        If Raza = HUMANO Then BonificadorRaza = 1
        If Raza = ElfoOscuro Then BonificadorRaza = 2
        If Raza = Enano Then BonificadorRaza = 3
        If Raza = ELFO Then BonificadorRaza = 0
        If Raza = Orco Then BonificadorRaza = 5
        If Raza = Gnomo Then BonificadorRaza = -5
    Case Agilidad
        If Raza = HUMANO Then BonificadorRaza = 1
        If Raza = ElfoOscuro Then BonificadorRaza = 0
        If Raza = Enano Then BonificadorRaza = -1
        If Raza = ELFO Then BonificadorRaza = 2
        If Raza = Orco Then BonificadorRaza = -2
        If Raza = Gnomo Then BonificadorRaza = 3
    Case Inteligencia
        If Raza = HUMANO Then BonificadorRaza = 1
        If Raza = ElfoOscuro Then BonificadorRaza = 2
        If Raza = Enano Then BonificadorRaza = -7
        If Raza = ELFO Then BonificadorRaza = 3
        If Raza = Orco Then BonificadorRaza = -6
        If Raza = Gnomo Then BonificadorRaza = 4
    Case Carisma
        If Raza = HUMANO Then BonificadorRaza = 0
        If Raza = ElfoOscuro Then BonificadorRaza = -1
        If Raza = Enano Then BonificadorRaza = -1
        If Raza = ELFO Then BonificadorRaza = 2
        If Raza = Orco Then BonificadorRaza = -4
        If Raza = Gnomo Then BonificadorRaza = 0
    Case Constitucion
        If Raza = HUMANO Then BonificadorRaza = 2
        If Raza = ElfoOscuro Then BonificadorRaza = 1
        If Raza = Enano Then BonificadorRaza = 4
        If Raza = ELFO Then BonificadorRaza = 0
        If Raza = Orco Then BonificadorRaza = 4
        If Raza = Gnomo Then BonificadorRaza = -1
End Select

End Function

Private Sub ResetAtributos()
lbAtributos.Caption = 40

Dim i As Integer

For i = 1 To NUMATRIBUTOS
    lbAtt(i - 1).Caption = "6"
    UserAtributos(i) = 6
Next i

End Sub
