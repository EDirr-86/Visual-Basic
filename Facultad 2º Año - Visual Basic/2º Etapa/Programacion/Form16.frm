VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   4560
      Width           =   3135
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label2 
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
      Left            =   4320
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Actual : "
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
      Left            =   3600
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6720
      Y1              =   4320
      Y2              =   4320
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

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.patch = Drive1.Drive
End Sub
Private Sub File1_Click()
   Label2.Caption = File1.List(File1.ListIndex)
End Sub

Private Sub File1_DblClick()
   MsgBox "Doble Click sobre " & File1.List(File1.ListIndex)
End Sub


