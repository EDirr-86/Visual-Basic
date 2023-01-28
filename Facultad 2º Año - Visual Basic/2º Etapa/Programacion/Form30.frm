VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Hexadecimal"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Octal"
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
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Resultado"
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
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Numero en decimal:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numero As Integer
Dim resultado As Integer
Dim residuo As Integer
Dim octa As Integer
Private Sub Option1_Click()
If Option1.Value = True Then
Text2 = Oct(Text1)
End If
End Sub
Private Sub Option2_Click()
If Option2.Value = True Then
Text2 = Hex(Text1)
End If
End Sub






