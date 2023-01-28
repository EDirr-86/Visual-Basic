VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cargar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Wi-Fi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   2535
      Begin VB.OptionButton Option7 
         Caption         =   "Ubuntu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   2175
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Mac"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.OptionButton Option4 
      Caption         =   "i7 Core"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   2775
   End
   Begin VB.OptionButton Option3 
      Caption         =   "i5 Core"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      Caption         =   "i3 Core"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Dual Core"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sist As String
Dim maquina As String

Private Sub Command1_Click()

If Option1.Value = True Then
    maquina = Option1.Caption
End If
If Option2.Value = True Then
    maquina = Option2.Caption
End If
If Option3.Value = True Then
    maquina = Option3.Caption
End If
If Option4.Value = True Then
    maquina = Option4.Caption
End If

If Option5.Value = True Then
    sist = Option5.Caption
End If
If Option6.Value = True Then
    sist = Option6.Caption
End If
If Option7.Value = True Then
    sist = Option7.Caption
End If

If Check1.Value = 0 Then
Text1.Text = " Selecciono " & maquina & " bajo el sistema operativo " & sist & " con conexion LAN"
Else
Text1.Text = " Selecciono " & maquina & " bajo el sistema operativo " & sist & " con conexion Wi-Fi"
End If
End Sub

Private Sub Command2_Click()
Text1 = ""
End Sub
