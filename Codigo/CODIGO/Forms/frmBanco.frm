VERSION 5.00
Begin VB.Form frmBanco 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Operación bancaria"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstBanco 
      Height          =   840
      ItemData        =   "frmBanco.frx":0000
      Left            =   120
      List            =   "frmBanco.frx":0010
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox txtDatos 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   4335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   345
      Left            =   3000
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmBanco.frx":0072
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblDatos 
      Caption         =   "¿Cuánto deseas depositar?"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   4335
   End
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Select Case lstBanco.ListIndex

    Case 0 'depositar oro
    
        'Si es negativo o cero jodete por pobre xD
        If Val(txtDatos.Text) <= 0 Then
            lblDatos.Caption = "Cantidad inválida."
            Exit Sub
        End If
        
    Case 1 'Retirar
    
        'Si es negativo o cero jodete por pobre xD
        If Val(txtDatos.Text) <= 0 Then
            lblDatos.Caption = "Cantidad inválida."
            Exit Sub
        End If
        
End Select

End Sub

Private Sub lstBanco_Click()

Select Case lstBanco.ListIndex
    Case 0 'Depositar oro
        lblDatos.Caption = "¿Cuánto deseas depositar?"
        txtDatos.Visible = True
    Case 1 'Retirar oro
        lblDatos.Caption = "¿Cuánto deseas retirar?"
        txtDatos.Visible = True
    Case 2 'ver la Boveda
        lblDatos.Caption = "Presiona Aceptra para ver tu Boveda."
        txtDatos.Visible = False
    Case 3 'Transferir oro
        lblDatos.Caption = "Completa los datos."
        txtDatos.Visible = False
End Select

End Sub
