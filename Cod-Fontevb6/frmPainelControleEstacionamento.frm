VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPainelControleEstacionamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Painel de Controle - Link"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9240
   Icon            =   "frmPainelControleEstacionamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2025
      Left            =   4485
      TabIndex        =   8
      Top             =   150
      Width           =   4650
      Begin VB.Label lblQtdeDentroEstacionamento 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2925
         TabIndex        =   14
         Top             =   1185
         Width           =   135
      End
      Begin VB.Label lblQtdeFilaSaida 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2925
         TabIndex        =   13
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lblQtdeFilaEntrada 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2925
         TabIndex        =   12
         Top             =   255
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dentro do Estacionamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   11
         Top             =   1170
         Width           =   2760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fila de Saída"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Top             =   690
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fila de Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1230
         TabIndex        =   9
         Top             =   225
         Width           =   1605
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Grade de Monitoramento"
      Height          =   5820
      Left            =   75
      TabIndex        =   6
      Top             =   2280
      Width           =   9090
      Begin MSFlexGridLib.MSFlexGrid griMonitoramento 
         Height          =   5340
         Left            =   90
         TabIndex        =   7
         Top             =   345
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   9419
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Controle por No. de Iterações"
      Height          =   1980
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   4275
      Begin VB.Timer timerProcessar 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3345
         Top             =   300
      End
      Begin VB.CommandButton cmsStartSistema 
         Caption         =   "Start Simulador"
         Height          =   510
         Left            =   2775
         TabIndex        =   5
         Top             =   915
         Width           =   1305
      End
      Begin VB.Frame Frame4 
         Caption         =   "Qtde Saídas"
         Height          =   480
         Left            =   210
         TabIndex        =   2
         Top             =   1260
         Width           =   2400
         Begin VB.TextBox txtQtdeSaidas 
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   4
            Text            =   "10"
            Top             =   15
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Qtde Entradas"
         Height          =   480
         Left            =   180
         TabIndex        =   1
         Top             =   525
         Width           =   2400
         Begin VB.TextBox txtQtdeEntradas 
            Height          =   300
            Left            =   1260
            MaxLength       =   8
            TabIndex        =   3
            Text            =   "10"
            Top             =   0
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "frmPainelControleEstacionamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================
'Simulador Controle de Estacionamento
'versão 1.0 teste
'10/08/2020
'Flávio R. Anselmi
'=======================================

Public aux, temp As Integer

Private Sub cmsStartSistema_Click()

    Call Carregar_gridMonitoramento _
            ( _
                Val(txtQtdeEntradas), _
                Val(txtQtdeSaidas) _
            )
    
    
    Call Iniciar_Processamento
    
End Sub

Private Sub Form_Load()

    Cabecalho_GridMonitoramento

End Sub

Private Sub timerProcessar_Timer()

    ContAuxTimer = ContAuxTimer + 1
    
    If ContAuxTimer > tempoAdecorrer Then
        timerProcessar.Enabled = False
        Call AtualizaStatus
    End If

End Sub

