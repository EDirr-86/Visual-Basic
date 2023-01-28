VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "&Titulo"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
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
      Height          =   435
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Sailr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&minuscula"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&MAYUSCULA"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox Text1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Texto convertido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese el texto a convertir:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2 = UCase(Text1)
End Sub

Private Sub Command2_Click()
Text2 = LCase(Text1)
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Text1 = ""
Text2 = ""
Text1.SetFocus
End Sub

Private Sub Command5_Click()
Text2.Text = UCase(Mid(Text1.Text, 1, 1)) + LCase(Mid(Text1, 2))
End Sub
