VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "OFF"
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
      Left            =   1680
      TabIndex        =   17
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   15
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   10
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   2160
      TabIndex        =   9
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   1200
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   2160
      TabIndex        =   6
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim primero As Double
Dim operacion As String
Private Sub Command1_Click(Index As Integer)
Text1.Text = Text1.Text + Command1(Index).Caption
End Sub

Private Sub Command2_Click()
primero = CDbl(Text1.Text)
Text1.Text = ""
operacion = "+"
End Sub

Private Sub Command3_Click()
primero = CDbl(Text1.Text)
Text1.Text = ""
operacion = "-"
End Sub

Private Sub Command4_Click()
primero = CDbl(Text1.Text)
Text1.Text = ""
operacion = "*"
End Sub

Private Sub Command5_Click()
primero = CDbl(Text1.Text)
Text1.Text = ""
operacion = "/"
End Sub

Private Sub Command6_Click()
Text1.Text = ""
End Sub
Private Sub Command7_Click()
If operacion = "+" Then
Text1.Text = primero + CDbl(Text1.Text)
End If

If operacion = "-" Then
Text1.Text = primero - CDbl(Text1.Text)
End If

If operacion = "*" Then
Text1.Text = primero * CDbl(Text1.Text)
End If

If operacion = "/" Then
Text1.Text = primero / CDbl(Text1.Text)
End If

End Sub

Private Sub Command8_Click()
End
End Sub
