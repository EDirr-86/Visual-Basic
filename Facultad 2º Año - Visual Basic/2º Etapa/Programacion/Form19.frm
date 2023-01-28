VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "---->"
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
      Left            =   2640
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "---->"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Grados Fahrenheit"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Grados Fahrenheit"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Grados Celsius"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Grados Celsius"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
Text3 = Val(Text1.Text) * 9 / 5 + 32
End Sub

Private Sub Text2_Change()
Text4 = (Val(Text2.Text) - 32) * 5 / 9
End Sub
