VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_clientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de clientes"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   13680
   Icon            =   "frm_clientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   13680
   Begin VB.Frame Frame5 
      Caption         =   "Pessoa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   59
      Top             =   0
      Width           =   3975
      Begin VB.OptionButton opt_fisica 
         Caption         =   "Física"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opt_juridica 
         Caption         =   "Jurídica"
         Height          =   195
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fra_buscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   7320
      TabIndex        =   29
      Top             =   480
      Width           =   6255
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txt_criterio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.ComboBox cbo_criterio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frm_clientes.frx":08CA
         Left            =   3120
         List            =   "frm_clientes.frx":08EF
         TabIndex        =   21
         Text            =   "( Todos os clientes )"
         Top             =   240
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frm_clientes.frx":093B
         TabIndex        =   30
         Top             =   360
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel lbl_criterio 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frm_clientes.frx":09D1
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSMask.MaskEdBox msk_criterio 
         Height          =   375
         Left            =   3120
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame fra_clientes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   7320
      TabIndex        =   28
      Top             =   2760
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid mfg_clientes 
         Height          =   6255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   11033
         _Version        =   393216
         Cols            =   4
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483636
         BackColorBkg    =   14737632
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   27
      Top             =   7440
      Width           =   7095
      Begin VB.Frame Frame6 
         Caption         =   "Ao expandir"
         Height          =   1575
         Left            =   4680
         TabIndex        =   60
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton Option2 
            Caption         =   "Verificar histórico"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   63
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Caption         =   "Buscar clientes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton cmd_expandir 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   61
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Manipulação de dados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Width           =   4455
         Begin VB.CommandButton cmd_novo 
            Caption         =   "Novo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmd_gravar 
            Caption         =   "Gravar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   57
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmd_alterar 
            Caption         =   "Alterar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   56
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmd_excluir 
            Caption         =   "Excluir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   55
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton cmd_validar 
         Cancel          =   -1  'True
         Caption         =   "Validar CPF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   53
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Endereço"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   26
      Top             =   4680
      Width           =   7095
      Begin VB.TextBox txt_bairro 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   16
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txt_complemento 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txt_logradouro 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   840
         Width           =   6135
      End
      Begin VB.TextBox txt_uf 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txt_cidade 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   19
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txt_numero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   1320
         Width           =   2175
      End
      Begin MSMask.MaskEdBox msk_cep 
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   51
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Compl."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   48
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Logra."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "CEP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Principais contatos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   7095
      Begin VB.TextBox txt_id 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   11
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txt_email 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   1320
         Width           =   6135
      End
      Begin MSMask.MaskEdBox msk_tel 
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cel 
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_nextel 
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "9999-9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   45
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Nextel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Celular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   42
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Fone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cliente / responsável"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   7095
      Begin VB.TextBox txt_nome 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   1320
         Width           =   6135
      End
      Begin VB.TextBox txt_ie 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txt_rsocial 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   6135
      End
      Begin MSMask.MaskEdBox msk_cnpj 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99.999.999/9999-99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_rg 
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   1800
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cpf 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "999.999.999/99"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "CPF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "RG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   39
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lbl_nome 
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "I.E."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   37
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "CNPJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl_rsocial 
         Caption         =   "R.Social"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.TextBox txt_codcli 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   20
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frm_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim L_Colunas, l_linha As Long
Dim L_codcli As Integer
Dim criterio As String
Dim valor As String
'''''''''''BD supermercados'''''''''''''''''''''''''''
Dim tabcli As New ADODB.Recordset 'tabela clientes
Dim tabcli2 As New ADODB.Recordset
Dim tabloca As New ADODB.Recordset 'Tabela localizações
Dim tabbairro As New ADODB.Recordset
Dim tabcid As New ADODB.Recordset
Dim tabuf As New ADODB.Recordset
Dim codigo As Integer
'''''''''''''''BD Ceps'''''''''''''''''''''''''''
Dim tab_loca As New ADODB.Recordset
Dim tab_bairro As New ADODB.Recordset
Dim tab_cid As New ADODB.Recordset
Dim tab_uf As New ADODB.Recordset
Dim cep As String
Dim logradouro As String
Dim bairro As String
Dim codbairro As String
Dim codcidade As String
Dim coduf As String
Option Explicit

Private Sub cbo_criterio_Click()
        Dim a As String
                dmascara
                msk_criterio = Empty
                txt_criterio = Empty
                hmascara
            If cbo_criterio <> "( Todos os clientes )" Then
                lbl_criterio.Visible = True
                lbl_criterio = cbo_criterio
                a = cbo_criterio
            ElseIf cbo_criterio = "( Todos os clientes )" Then
                    Call cmd_buscar_Click
                    lbl_criterio.Visible = False
                    txt_criterio.Visible = False
                    msk_criterio.Visible = False
            End If

            If a = "Código" Then
                txt_criterio.Visible = True
                msk_criterio.Visible = False
                txt_criterio.SetFocus
                
            ElseIf a = "RG" Then
                    txt_criterio.Visible = False
                    msk_criterio.Visible = True
                    msk_criterio.Mask = "99.999.999-&"
                    msk_criterio.SetFocus
                
               ' ElseIf a = "Nascimento" Then
                        'txt_criterio.Visible = False
                        'msk_criterio.Visible = True
                       '' msk_criterio.Mask = "99/99/9999"
                       ' msk_criterio.SetFocus
                        
                    ElseIf a = "Telefone" Then
                            txt_criterio.Visible = False
                            msk_criterio.Visible = True
                            msk_criterio.Mask = "(99)9999-9999"
                            msk_criterio.SetFocus
                            
                        ElseIf a = "Celular" Then
                                txt_criterio.Visible = False
                                msk_criterio.Visible = True
                                msk_criterio.Mask = "(99)9999-9999"
                                msk_criterio.SetFocus
                                
                            ElseIf a = "Cep" Then
                                    txt_criterio.Visible = False
                                    msk_criterio.Visible = True
                                    msk_criterio.Mask = "99999-999"
                                    msk_criterio.SetFocus
                                
                                ElseIf a = "CPF" Then
                                        txt_criterio.Visible = False
                                        msk_criterio.Visible = True
                                        msk_criterio.Mask = "999.999.999/99"
                                        msk_criterio.SetFocus
                                    
                                    ElseIf a = "Nome" Then
                                            txt_criterio.Visible = True
                                            msk_criterio.Visible = False
            End If



End Sub


Private Sub cmd_alterar_Click()
            status = "alteradas"
            tabcli.Close
            tabcli.Open "select * from clientes where codigo like '" & txt_codcli & "'"
            If tabcli.RecordCount <> 0 Then
                Call gravar
                Call box
                Call codcli
                If cbo_criterio = "( Todos os clientes )" Then
                    Call carregar_lista
                End If
            End If
End Sub

Private Sub cmd_buscar_Click()
            dmascara
                criterio = cbo_criterio
            If criterio = "Código" Then
                criterio = "Codigo"
                valor = txt_criterio
            ElseIf criterio = "Nome" Then
                    valor = txt_criterio
                ElseIf criterio = "RG" Then
                        criterio = "Rg"
                        valor = msk_criterio
                    ElseIf criterio = "Celular" Then
                            valor = msk_criterio
                        ElseIf criterio = "Telefone" Then
                                criterio = "Tel_res"
                                valor = msk_criterio
                            ElseIf criterio = "Cep" Then
                                    valor = msk_criterio
                                ElseIf criterio = "CPF" Then
                                        criterio = "Cpf"
                                        valor = msk_criterio
            End If
            If cbo_criterio = "( Todos os clientes )" Then
                Call carregar_lista
            ElseIf criterio <> "( Todos os clientes )" Then
                Call clcc
            End If
            dmascara
End Sub

Private Sub cmd_excluir_Click()
            status = "excluidas"
            tabcli.Close
            tabcli.Open "select * from Clientes where Codigo like '" & txt_codcli & "'"
            If MsgBox("Deseja realmente excluir este cliente?", vbYesNo + vbDefaultButton2 + vbQuestion, "Arbimy Manager 2.0") = vbYes Then
            If tabcli.RecordCount <> 0 Then
                conectar.Execute "delete from Clientes where Codigo like '" & txt_codcli & "'"
            Call cmd_novo_Click
            Call box
            Call codcli
            If cbo_criterio = "( Todos os clientes )" Then
                Call carregar_lista
            End If
            End If
            End If
End Sub

Private Sub cmd_expandir_Click()
            
            If frm_clientes.Width = 7395 Then
                frm_clientes.ScaleWidth = 13680
                frm_clientes.Width = 13770
                cmd_expandir.Caption = "<<"
            ElseIf frm_clientes.Width = 13770 Then
                    Call dimensoes
                    cmd_expandir.Caption = ">>"
            End If
            
End Sub
Private Sub cmd_gravar_Click()
            status = "gravadas"
            Call dmascara
            
            tabcli.Close
            tabcli.Open "select * from clientes where Rg like '" & msk_rg & "'"
                If tabcli.RecordCount = 1 Then
                    MsgBox "Este Rg já está cadastrado, favor verificar", vbInformation, "Arbimy Manager 2.0"
                    msk_rg = Empty
                    msk_rg.SetFocus
                    Exit Sub
                ElseIf tabcli.RecordCount = 0 Then
                        tabcli.Close
                        tabcli.Open "select * from clientes where Cpf like '" & msk_cpf & "'"
                            If tabcli.RecordCount = 1 Then
                                MsgBox "Este Cpf já está cadastrado, favor verificar", vbInformation, "Arbimy Manager 2.0"
                                msk_cpf = Empty
                                msk_cpf.SetFocus
                                Exit Sub
                            End If
                End If
            
            If txt_nome = Empty Then
                txt_nome.BackColor = &HFF&
                txt_nome.SetFocus
                MsgBox "Existem informações que não foram preenchidas, favor verificar", vbInformation, "Arbimy Manager 2.0"
                Exit Sub
            Else
            If msk_cep = Empty Then
                msk_cep.BackColor = &HFF&
                MsgBox "Existem informações que não foram preenchidas, favor verificar", vbInformation, "Arbimy Manager 2.0"
                Exit Sub
            End If
            End If
            
            
            Call gravar_loca
            Call abrir
            Call gravar
            Call box
            Call codcli
                If cbo_criterio = "( Todos os clientes )" Then
                    Call carregar_lista
                End If
            Call hmascara
End Sub

Private Sub cmd_novo_Click()
            Call dmascara
                txt_codcli = Empty
                txt_nome = Empty
                txt_email = Empty
                txt_numero = Empty
                txt_cidade = Empty
                txt_uf = Empty
                txt_logradouro = Empty
                txt_criterio = Empty
                txt_complemento = Empty
                txt_bairro = Empty
                 msk_cep = Empty
                 msk_rg = Empty
                 msk_cpf = Empty
                 msk_tel = Empty
                 msk_cel = Empty
                 msk_criterio = Empty
                'dtp_nascimento = Date
                cbo_criterio = "( Todos os clientes )"
            txt_nome.SetFocus
            Call hmascara
            Call codcli
End Sub



Private Sub cmd_validar_Click()
           msk_cpf.PromptInclude = False
                If msk_cpf.Text = Empty Then Exit Sub
            CPF = msk_cpf.Text
            strCampo = Left(CPF, 9)
            Call CalculaCPF
            If StrConf <> Right(CPF, 2) Then
               MsgBox "Número do CPF Inválido. Por favor, tente novamente.", vbInformation, "Arbimy Manager 2.0"
                       msk_cpf = Empty
                       msk_cpf.PromptInclude = True
                       msk_cpf.SetFocus
                Else
                    MsgBox "Cpf Válido", vbInformation
            End If
End Sub

Private Sub Form_Load()
            Call abrir_banco
            'dtp_nascimento.Value = Date
            Call dimensoes
            
            Call abrir
            Call codcli
            Call carregar_lista
End Sub
Private Sub dimensoes()
'            frm_clientes.Height = 6975
'            frm_clientes.ScaleHeight = 6510
            frm_clientes.ScaleWidth = 7305
            frm_clientes.Width = 7395
End Sub

Private Sub mfg_clientes_Click()
            l_linha = mfg_clientes.Row
            L_codcli = mfg_clientes.TextMatrix(l_linha, 0)
                Call abrir
            tabcli.Close
            tabcli.Open "Select * From Clientes Where Codigo = " & L_codcli
            Call mostrar
End Sub

Private Sub msk_cel_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub msk_cep_GotFocus()
            If msk_cep.BackColor <> &H80000005 Then
                msk_cep.BackColor = &H80000005
            End If
End Sub

Private Sub msk_cep_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub msk_cep_LostFocus()
            dmascara
                If msk_cep = Empty Then
                    Exit Sub
                End If
                cep = msk_cep
                    Call logra
                Call bair
                Call cid
                Call uf
            hmascara
End Sub

Private Sub msk_cpf_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub validar()
           
End Sub

Private Sub msk_criterio_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub msk_rg_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub msk_tel_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub SkinLabel1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub opt_fisica_Click()
            If opt_fisica.Enabled Then
                txt_rsocial.Enabled = False
                msk_cnpj.Enabled = False
                txt_ie.Enabled = False
                txt_nome.Enabled = True
                msk_cpf.Enabled = True
                msk_rg.Enabled = True
                    txt_nome.SetFocus
                    
            End If
            
End Sub

Private Sub opt_juridica_Click()
            If opt_juridica.Enabled Then
                    txt_rsocial.Enabled = True
                msk_cnpj.Enabled = True
                txt_ie.Enabled = True
                txt_nome.Enabled = False
                msk_cpf.Enabled = False
                msk_rg.Enabled = False
                    txt_rsocial.SetFocus
            End If
End Sub

Private Sub txt_complemento_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub txt_criterio_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub txt_email_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub txt_nome_Change()
            If txt_nome.BackColor = &HFF& Then
                txt_nome.BackColor = &HFFFFFF
            End If
End Sub

Private Sub txt_nome_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
            
            If KeyAscii = 8 Then
                KeyAscii = 8
            ElseIf KeyAscii = 32 Then
                    KeyAscii = 32
                    
                ElseIf KeyAscii < 65 Or KeyAscii > 90 Then
                    If KeyAscii < 97 Or KeyAscii > 123 Then
                        KeyAscii = 0
            End If
            End If
End Sub

Private Sub txt_nome_LostFocus()
            txt_nome = UCase(txt_nome)
End Sub
Private Sub fechar() ' fechar tabelas do BD Supermercados
            If tabcli.State = 1 Then tabcli.Close
            If tabcli2.State = 1 Then tabcli2.Close
            If tabloca.State = 1 Then tabloca.Close
            If tabbairro.State = 1 Then tabbairro.Close
            If tabcid.State = 1 Then tabcid.Close
            If tabuf.State = 1 Then tabuf.Close
End Sub
Private Sub abrir() ' abrir tabelas do BD Supermercados
            Call fechar
            tabcli.Open "Clientes", conectar, adOpenKeyset, adLockOptimistic
            tabcli2.Open "Clientes", conectar, adOpenKeyset, adLockOptimistic
            tabloca.Open "Localizacoes", conectar, adOpenKeyset, adLockOptimistic
            tabbairro.Open "Bairros", conectar, adOpenKeyset, adLockOptimistic
            tabcid.Open "Cidades", conectar, adOpenKeyset, adLockOptimistic
            tabuf.Open "Ufs", conectar, adOpenKeyset, adLockOptimistic
End Sub
Private Sub codcli()
            codigo = 1
a:
            tabcli2.Close
            tabcli2.Open "select * from clientes where Codigo like '" & codigo & "'"
            If tabcli2.RecordCount = 1 Then
            codigo = codigo + 1
            GoTo a:
            End If
            txt_codcli = codigo
End Sub
Private Sub dmascara()
            msk_cep.PromptInclude = False
            msk_rg.PromptInclude = False
            msk_cpf.PromptInclude = False
            msk_tel.PromptInclude = False
            msk_cel.PromptInclude = False
            msk_criterio.PromptInclude = False
End Sub
Private Sub hmascara()
            msk_cep.PromptInclude = True
            msk_rg.PromptInclude = True
            msk_cpf.PromptInclude = True
            msk_tel.PromptInclude = True
            msk_cel.PromptInclude = True
            msk_criterio.PromptInclude = True
End Sub
Private Sub gravar()
            Call dmascara
            If status <> "alteradas" Then
                tabcli.Close
                tabcli.Open "Select * from clientes where codigo like '" & txt_codcli & "' "
                If tabcli.RecordCount = 0 Then
                tabcli.AddNew
            End If
            End If
                tabcli!nome = txt_nome
                tabcli!cep = msk_cep
                tabcli!numero = txt_numero
                tabcli!Complemento = txt_complemento
               ' tabcli!Nascimento = dtp_nascimento.Value
                tabcli!Tel_res = msk_tel
                tabcli!Celular = msk_cel
                tabcli!codigo = txt_codcli
                tabcli!RG = msk_rg
                tabcli!Email = txt_email
                tabcli!CPF = msk_cpf
            tabcli.Update
            If status = "gravadas" Then
                Call cmd_novo_Click
            End If
            Call hmascara
End Sub
Private Sub logra()

            Call abrir_banco2
            Call abrirc
                tab_loca.Close
                tab_loca.Open "Select * from Endereco where Endereco_CEP = '" & msk_cep & "'"
            If tab_loca.RecordCount = 1 Then
                logradouro = tab_loca!Endereco_Logradouro
                txt_logradouro = logradouro
                    codbairro = tab_loca!Bairro_Codigo
                    codcidade = tab_loca!Cidade_Codigo
                    coduf = tab_loca!UF_Codigo
            ElseIf tab_loca.RecordCount = 0 Then
                    MsgBox "Código de Endereçamento Postal (CEP) não encontrado, favor verificar", vbExclamation, "Arbimy Manager 2.0"
                        msk_cep = Empty
                        msk_cep.SetFocus
            End If
End Sub
Private Sub fecharc() ' fechar tabelas do BD Ceps
            If tab_loca.State = 1 Then tab_loca.Close
            If tab_bairro.State = 1 Then tab_bairro.Close
            If tab_cid.State = 1 Then tab_cid.Close
            If tab_uf.State = 1 Then tab_uf.Close
End Sub
Private Sub abrirc() ' abrir tabelas do BD Ceps
            Call fecharc
            tab_loca.Open "Endereco", conectar2, adOpenKeyset, adLockOptimistic
            tab_bairro.Open "Bairro", conectar2, adOpenKeyset, adLockOptimistic
            tab_cid.Open "Cidade", conectar2, adOpenKeyset, adLockOptimistic
            tab_uf.Open "UF", conectar2, adOpenKeyset, adLockOptimistic
End Sub
Private Sub bair()
            tab_bairro.Close
            tab_bairro.Open "select * from Bairro where Bairro_Codigo like '" & codbairro & "'"
            If tab_bairro.RecordCount = 1 Then
                    bairro = tab_bairro!Bairro_Descricao
                    txt_bairro = bairro
            End If
End Sub
Private Sub cid()
            tab_cid.Close
            tab_cid.Open "select * from Cidade where Cidade_Codigo like '" & codcidade & "'"
            If tab_cid.RecordCount = 1 Then
                txt_cidade = tab_cid!Cidade_Descricao
            End If
End Sub
Private Sub uf()
            tab_uf.Close
            tab_uf.Open "select * from UF where uf_codigo like '" & coduf & "'"
            If tab_uf.RecordCount = 1 Then
                txt_uf = tab_uf!uf_sigla
            End If
End Sub
Private Sub gravar_loca()
            Call abrir
            dmascara
            tabuf.Close
            tabuf.Open "select * from Ufs where Codigo like '" & coduf & "'"
            If tabuf.RecordCount = 0 Then
                tabuf.AddNew
                tabuf!codigo = coduf
                tabuf!Estado = txt_uf
                tabuf.Update
            End If
                tabcid.Close
                tabcid.Open "select * from Cidades where cod_cidade like '" & codcidade & "'"
                If tabcid.RecordCount = 0 Then
                    tabcid.AddNew
                    tabcid!Cod_cidade = codcidade
                    tabcid!Cidade = txt_cidade
                    tabcid!Cod_estado = coduf
                    tabcid.Update
                End If
                    tabbairro.Close
                    tabbairro.Open "select * from Bairros where Cod_Bairro like '" & codbairro & "'"
                    If tabbairro.RecordCount = 0 Then
                        tabbairro.AddNew
                        tabbairro!Cod_bairro = codbairro
                        tabbairro!bairro = txt_bairro
                        tabbairro!Cod_cidade = codcidade
                        tabbairro.Update
                    End If
                        tabloca.Close
                        tabloca.Open "select * from Localizacoes where Cep = '" & msk_cep & "'"
                        If tabloca.RecordCount = 0 Then
                            tabloca.AddNew
                            tabloca!cep = msk_cep
                            tabloca!logradouro = txt_logradouro
                            tabloca!Cod_bairro = codbairro
                            tabloca.Update
                        End If
            hmascara
End Sub
Private Sub carregar_lista()
            Call abrir
            If tabcli.BOF = False Or tabcli.EOF = False Then
                tabcli.MoveFirst
                    mfg_clientes.Rows = 2
                    mfg_clientes.Clear
                    mfg_clientes.FormatString = "Código  |Nome                                                         |RG                   |Cep            "
                Do Until tabcli.EOF
'                    If (mfg_clientes.Row Mod 2) = 0 Then
'                        mfg_clientes.CellBackColor = vbBlue
'
'                    Else
'                        mfg_clientes.CellForeColor = vbRed
'                    End If
                    
                    mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 0) = tabcli!codigo
                    mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 1) = tabcli!nome
                    mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 2) = Format(tabcli!RG, "&&.&&&.&&&-&")
                    mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 3) = Format(tabcli!cep, "&&&&&-&&&")
                        mfg_clientes.Rows = mfg_clientes.Rows + 1
'                        mfg_clientes.Row = mfg_clientes.Row + 1
                    tabcli.MoveNext
                Loop
                     
                    mfg_clientes.Rows = mfg_clientes.Rows - 1
                Else
                    mfg_clientes.Rows = 2
                    mfg_clientes.Clear
                    mfg_clientes.FormatString = "Código  |Nome                                  |RG                           |Cep                        "
            End If
            
            
End Sub
Private Sub mostrar()
            dmascara
            txt_codcli = tabcli!codigo
            txt_nome = tabcli!nome
            txt_numero = tabcli!numero
            txt_complemento = tabcli!Complemento
            txt_email = tabcli!Email
                msk_rg = tabcli!RG
                msk_cpf = tabcli!CPF
                msk_tel = tabcli!Tel_res
                msk_cel = tabcli!Celular
                msk_cep = tabcli!cep
            ' dtp_nascimento.Value = tabcli!Nascimento
             Call msk_cep_LostFocus
            hmascara
End Sub
Private Sub clcc()
                Call abrir
                tabcli.Close
                tabcli.Open "Select * from Clientes where " & criterio & " like '" & valor & "'"
                If tabcli.RecordCount = 0 Then
                    MsgBox "Não Existem clientes que possuem este(a) " & cbo_criterio & "", vbInformation, "Arbimy Manager 2.0"
                            Call cbo_criterio_Click
                                If msk_criterio.Visible = True Then
                                    msk_criterio.SetFocus
                                ElseIf txt_criterio.Visible = True Then
                                        txt_criterio.SetFocus
                                End If
                            Exit Sub
                ElseIf tabcli.RecordCount > 0 Then
                tabcli.MoveFirst
                    mfg_clientes.Rows = 2
                    mfg_clientes.Clear
                    mfg_clientes.FormatString = "Código  |Nome                                                         |RG                   |Cep            "
                Do Until tabcli.EOF
                    mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 0) = tabcli!codigo
                    mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 1) = tabcli!nome
                    mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 2) = Format(tabcli!RG, "&&.&&&.&&&-&")
                    mfg_clientes.TextMatrix(mfg_clientes.Rows - 1, 3) = Format(tabcli!cep, "&&&&&-&&&")

                        mfg_clientes.Rows = mfg_clientes.Rows + 1
                    tabcli.MoveNext
                Loop
                    mfg_clientes.Rows = mfg_clientes.Rows - 1
                Else
                    mfg_clientes.Rows = 2
                    mfg_clientes.Clear
                    mfg_clientes.FormatString = "Código  |Nome                                  |RG                           |Cep                        "
            End If
End Sub

Private Sub txt_numero_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
End Sub

Private Sub txt_rsocial_LostFocus()
            txt_rsocial = UCase(txt_rsocial)
End Sub
