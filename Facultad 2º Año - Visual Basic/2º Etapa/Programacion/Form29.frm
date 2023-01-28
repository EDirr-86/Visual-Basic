VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   3180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Calcular"
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
      TabIndex        =   9
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
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
      Left            =   240
      TabIndex        =   8
      Top             =   4440
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      MaxLength       =   10
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "dd/mm/aaa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Años:"
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
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Mes:"
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
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Dias:"
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
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fecha de nacimiento"
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
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   2055
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
Text2 = DateDiff("d", Text1, Now)
Text3 = DateDiff("m", Text1, Now)
If Val(Mid(Text1, 1, 5)) = Val(Mid(Now, 1, 5)) Then
Text4 = DateDiff("yyyy", Text1, Now)
Else
Text4 = DateDiff("yyyy", Text1, Now) - 1
End If
End Sub

Private Sub Command3_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub
