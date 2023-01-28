VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   6405
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
      Left            =   2280
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Zoom"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   360
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   3360
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Image1.Picture = Picture1.Picture
Form2.Show 1
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.patch = Drive1.Drive
End Sub
Private Sub File1_Click()
Picture1.Picture = LoadPicture(App.Path + File1.List(File1.ListIndex)) 'propiedad pattern filtra *.*'
End Sub
Private Sub File1_DblClick()
   MsgBox "Doble Click sobre " & File1.List(File1.ListIndex)
End Sub


