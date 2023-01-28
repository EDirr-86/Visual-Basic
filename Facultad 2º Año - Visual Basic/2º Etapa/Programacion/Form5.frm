VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "&Salir"
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
      Left            =   1920
      TabIndex        =   9
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formulario 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.CommandButton Command4 
         Caption         =   "Ocultar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Descargar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cargar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Metodo .Hide"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Metodo .Show"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   " Sentencia Unload"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Sentencia Load"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   840
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form2
MsgBox "Se cargo en memoria"
End Sub

Private Sub Command2_Click()
Unload Form2
MsgBox "Se descargo de memoria"
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub

Private Sub Command4_Click()
Form2.Hide
End Sub

Private Sub Command5_Click()
End
End Sub
