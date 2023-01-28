VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Calcular"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calculo de la hipotenusa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   7
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2160
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Cateto 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Cateto 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Text3 = Sqr(Val(Text1 * Text1) + Val(Text2 * Text2))
End Sub
