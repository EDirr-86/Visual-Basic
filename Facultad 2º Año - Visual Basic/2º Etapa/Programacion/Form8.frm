VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   4635
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cagar"
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
      Top             =   2760
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   2520
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2460
      Left            =   240
      Picture         =   "Form8.frx":0000
      ScaleHeight     =   2400
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture2.Picture = LoadPicture(App.Path + "Sk.jpg")
End Sub
Private Sub Command2_Click()
End
End Sub
