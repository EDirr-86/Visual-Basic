VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "SUMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   2160
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Pares"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Impares"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim impar As Integer
Dim suma As Integer
Dim par As Integer
Private Sub Command1_Click()
suma = 0
Text1 = ""
For impar = 1 To 99 Step 2
      suma = suma + impar
Next
Text1 = suma
End Sub

Private Sub Command2_Click()
suma = 0
Text1 = ""
For impar = 0 To 98 Step 2
      suma = suma + impar
Next
Text1 = suma
End Sub
