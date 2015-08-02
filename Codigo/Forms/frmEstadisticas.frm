VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   ClipControls    =   0   'False
   Icon            =   "frmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEstadisticas.frx":000C
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image command1 
      Height          =   105
      Index           =   33
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":42EE0
      Top             =   960
      Width           =   195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   6240
      TabIndex        =   44
      Top             =   120
      Width           =   135
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Criaturas matadas"
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
      Height          =   195
      Index           =   1
      Left            =   2250
      TabIndex        =   43
      Top             =   6180
      Width           =   1665
   End
   Begin VB.Label est 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Veces muerto"
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
      Height          =   195
      Index           =   4
      Left            =   2280
      TabIndex        =   42
      Top             =   5550
      Width           =   1665
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   6
      Left            =   1230
      TabIndex        =   41
      Top             =   5700
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   5
      Left            =   1230
      TabIndex        =   40
      Top             =   5490
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   4
      Left            =   1230
      TabIndex        =   39
      Top             =   5250
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   3
      Left            =   1230
      TabIndex        =   38
      Top             =   4890
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   2
      Left            =   1200
      TabIndex        =   37
      Top             =   4680
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00000"
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
      Height          =   165
      Index           =   1
      Left            =   1230
      TabIndex        =   36
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Clase"
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
      Height          =   135
      Index           =   0
      Left            =   960
      TabIndex        =   35
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Género"
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
      Height          =   135
      Index           =   5
      Left            =   960
      TabIndex        =   34
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label est 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Raza"
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
      Height          =   135
      Index           =   6
      Left            =   960
      TabIndex        =   33
      Top             =   3330
      Width           =   975
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   1
      Left            =   1590
      TabIndex        =   32
      Top             =   720
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   2
      Left            =   1590
      TabIndex        =   31
      Top             =   975
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   3
      Left            =   1590
      TabIndex        =   30
      Top             =   1260
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   4
      Left            =   1590
      TabIndex        =   29
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label Atri 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   165
      Index           =   5
      Left            =   1590
      TabIndex        =   28
      Top             =   1800
      Width           =   105
   End
   Begin VB.Image cmdGuardar 
      Height          =   360
      Left            =   3840
      Tag             =   "1"
      Top             =   3960
      Width           =   930
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   195
      Index           =   26
      Left            =   5740
      TabIndex        =   27
      Top             =   3400
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   5740
      TabIndex        =   26
      Top             =   3180
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   5740
      TabIndex        =   25
      Top             =   2960
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   5740
      TabIndex        =   24
      Top             =   2720
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   5740
      TabIndex        =   23
      Top             =   2500
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   5740
      TabIndex        =   22
      Top             =   2280
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   5740
      TabIndex        =   21
      Top             =   1600
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   5740
      TabIndex        =   20
      Top             =   1820
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   5740
      TabIndex        =   19
      Top             =   2040
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   5740
      TabIndex        =   18
      Top             =   1380
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   5745
      TabIndex        =   17
      Top             =   1150
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   5745
      TabIndex        =   16
      Top             =   930
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   5740
      TabIndex        =   15
      Top             =   710
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   4050
      TabIndex        =   14
      Top             =   3620
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   4050
      TabIndex        =   13
      Top             =   3400
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   4050
      TabIndex        =   12
      Top             =   3180
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   4050
      TabIndex        =   11
      Top             =   2960
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4050
      TabIndex        =   10
      Top             =   2720
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4050
      TabIndex        =   9
      Top             =   2500
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4050
      TabIndex        =   8
      Top             =   2280
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4050
      TabIndex        =   7
      Top             =   2040
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4050
      TabIndex        =   6
      Top             =   1810
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4050
      TabIndex        =   5
      Top             =   1590
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4050
      TabIndex        =   4
      Top             =   1370
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4050
      TabIndex        =   3
      Top             =   1150
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4050
      TabIndex        =   2
      Top             =   710
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4050
      TabIndex        =   1
      Top             =   930
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   3
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":43032
      Top             =   1020
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   5
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":43184
      Top             =   1260
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   7
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":432D6
      Top             =   1500
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   9
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":43428
      Top             =   1740
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   11
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":4357A
      Top             =   1950
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   13
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":436CC
      Top             =   2190
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   15
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":4381E
      Top             =   2400
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   17
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":43970
      Top             =   2610
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   19
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":43AC2
      Top             =   2850
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   21
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":43C14
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   23
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":43D66
      Top             =   3270
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   25
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":43EB8
      Top             =   3510
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   27
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":4400A
      Top             =   3750
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   1
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":4415C
      Top             =   810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   0
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":442AE
      Top             =   720
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   2
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":44400
      Top             =   930
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   4
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":44552
      Top             =   1170
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   6
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":446A4
      Top             =   1380
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   8
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":447F6
      Top             =   1620
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   10
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":44948
      Top             =   1860
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   12
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":44A9A
      Top             =   2070
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   14
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":44BEC
      Top             =   2310
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   16
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":44D3E
      Top             =   2520
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   18
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":44E90
      Top             =   2760
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   20
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":44FE2
      Top             =   2970
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   22
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":45134
      Top             =   3180
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   24
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":45286
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   26
      Left            =   4310
      MouseIcon       =   "frmEstadisticas.frx":453D8
      Top             =   3660
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   53
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":4552A
      Top             =   3510
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   52
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":4567C
      Top             =   3420
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   51
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":457CE
      Top             =   3300
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   50
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":45920
      Top             =   3210
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   49
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":45A72
      Top             =   3060
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   48
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":45BC4
      Top             =   2970
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   47
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":45D16
      Top             =   2850
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   46
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":45E68
      Top             =   2760
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   45
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":45FBA
      Top             =   2640
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   44
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":4610C
      Top             =   2520
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   43
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":4625E
      Top             =   2430
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   42
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":463B0
      Top             =   2310
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   28
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":46502
      Top             =   720
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   29
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":46654
      Top             =   810
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   30
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":467A6
      Top             =   960
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   31
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":468F8
      Top             =   1320
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   34
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":46A4A
      Top             =   1410
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   35
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":46B9C
      Top             =   1500
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":46CEE
      Top             =   1620
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   37
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":46E40
      Top             =   1710
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   38
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":46F92
      Top             =   1830
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   39
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":470E4
      Top             =   1950
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   40
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":47236
      Top             =   2070
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   41
      Left            =   6000
      MouseIcon       =   "frmEstadisticas.frx":47388
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   195
      Left            =   5880
      TabIndex        =   0
      Top             =   3800
      Width           =   285
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
Private LibresOrig As Integer
Private RealizoCambios As Boolean
Private NewSkills(1 To NUMSKILLS) As Byte
Private Sub cmdGuardar_Click()
Dim i As Byte
If RealizoCambios Then
    If MsgBox("Los skills han sido modificados. Deseas guardarlos?", vbYesNo) = vbYes Then
        For i = 1 To NUMSKILLS
            NewSkills(i) = CByte(Skill(i - 1).Caption) - UserSkills(i)
            UserSkills(i) = Val(Skill(i - 1).Caption)
        Next i
        
        Call WriteModifySkills(NewSkills())
    End If
End If
Unload Me
End Sub


Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim i As Integer
For i = 1 To NUMATRIBUTOS
    Atri(i).Caption = UserAtributos(i)
Next

For i = 1 To NUMSKILLS
    Skill(i - 1).Caption = UserSkills(i)
Next

With UserEstadisticas
    Label4(1).Caption = .CiudadanosMatados
    Label4(2).Caption = .RepublicanosMatados
    Label4(3).Caption = .RenegadosMatados
    Label4(4).Caption = .ArmadaMatados
    Label4(5).Caption = .MiliciaMatados
    Label4(6).Caption = .CaosMatados
    
    est(0).Caption = ListaClases(.Clase)
    est(5).Caption = IIf(.Genero = 1, "Masculino", "Femenino")
    est(6).Caption = ListaRazas(.Raza)
End With
LibresOrig = SkillPoints

Puntos.Caption = SkillPoints
RealizoCambios = False

End Sub

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
    indice = Index \ 2
    If Val(Skill(indice).Caption) > 0 And Not (Val(Skill(indice).Caption) = SkillsOrig(indice + 1)) Then
        Skill(indice).Caption = Val(Skill(indice).Caption) - 1
        SkillPoints = SkillPoints + 1
    End If
End If

Puntos.Caption = SkillPoints
RealizoCambios = (SkillPoints <> LibresOrig)
Skill(indice).ForeColor = IIf(Val(Skill(indice).Caption) = SkillsOrig(indice + 1), vbWhite, vbRed)

End Sub

Private Sub Form_Load()
    Me.Picture = LoadInterface("estisticas")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub s_Click()

End Sub

Public Function SkillRealToIndex(ByVal SkillIndex As Integer) As Integer

Select Case SkillIndex
    Case 1
        SkillRealToIndex = 4
    Case 2
        SkillRealToIndex = 5
    Case 3
        SkillRealToIndex = 20
    Case 4
        SkillRealToIndex = 7
    Case 5
        SkillRealToIndex = 23
    Case 6
        SkillRealToIndex = 19
    Case 7
        SkillRealToIndex = 12
    Case 8
        SkillRealToIndex = 2
    Case 9
        SkillRealToIndex = 22
    Case 10
        SkillRealToIndex = 6
    Case 11
        SkillRealToIndex = 8
    Case 12
        SkillRealToIndex = 18
    Case 13
        SkillRealToIndex = 1
    Case 14
        SkillRealToIndex = 3
    Case 15
        SkillRealToIndex = 11
    Case 16
        SkillRealToIndex = 9
    Case 17
        SkillRealToIndex = 17
    Case 18
        SkillRealToIndex = 13
    Case 19
        SkillRealToIndex = 14
    Case 20
        SkillRealToIndex = 10
    Case 21
        SkillRealToIndex = 26
    Case 22
        SkillRealToIndex = 16
    Case 23
        SkillRealToIndex = 15
    Case 24
        SkillRealToIndex = 24
    Case 25
        SkillRealToIndex = 25
    Case 26
        SkillRealToIndex = 21
    Case 27
        SkillRealToIndex = 27
End Select

End Function
Public Function RealSkillToIndex(ByVal Skill As Integer) As Integer

Select Case Skill
    Case 4
        RealSkillToIndex = 1
    Case 5
        RealSkillToIndex = 2
    Case 20
        RealSkillToIndex = 3
    Case 7
        RealSkillToIndex = 4
    Case 23
        RealSkillToIndex = 5
    Case 19
        RealSkillToIndex = 6
    Case 12
        RealSkillToIndex = 7
    Case 2
        RealSkillToIndex = 8
    Case 22
        RealSkillToIndex = 9
    Case 6
        RealSkillToIndex = 10
    Case 8
        RealSkillToIndex = 11
    Case 18
        RealSkillToIndex = 12
    Case 1
        RealSkillToIndex = 13
    Case 3
        RealSkillToIndex = 14
    Case 11
        RealSkillToIndex = 15
    Case 9
        RealSkillToIndex = 16
    Case 17
        RealSkillToIndex = 17
    Case 13
        RealSkillToIndex = 18
    Case 14
        RealSkillToIndex = 19
    Case 10
        RealSkillToIndex = 20
    Case 26
        RealSkillToIndex = 21
    Case 16
        RealSkillToIndex = 22
    Case 15
        RealSkillToIndex = 23
    Case 24
        RealSkillToIndex = 24
    Case 25
        RealSkillToIndex = 25
    Case 21
        RealSkillToIndex = 26
    Case 27
        RealSkillToIndex = 27
End Select

End Function

Private Sub Label1_Click()
Unload Me
End Sub
