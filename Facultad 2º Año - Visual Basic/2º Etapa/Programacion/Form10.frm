VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
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
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
      Begin VB.OptionButton Option4 
         Caption         =   "Club Atletico River Plate"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Club Atletico Boca Juniors"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sexo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton Option2 
         Caption         =   "Femenino"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Masculino"
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
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value Then
    MsgBox "Sexo: " & Option1.Caption
Else
    MsgBox "Sexo: " & Option2.Caption
End If

If Option3.Value Then
    MsgBox "Club: " & Option3.Caption
Else
    MsgBox "Club: " & Option4.Caption
End If

End Sub

Private Sub Command2_Click()
End
End Sub
