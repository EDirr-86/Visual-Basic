VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4920
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
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Timer Timer2 
      Left            =   2520
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Activar"
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
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
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
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Despertador"
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
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Hora Actual"
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
      TabIndex        =   0
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
If Command1.Caption = "&Activar" Then
Timer2.Enabled = True
Timer2.Interval = 1000
Command1.Caption = "&Desactivar"
Else
Timer2.Enabled = False
Command1.Caption = "&Activar"
Text1.Text = ""
End If
End Sub
Private Sub Command2_Click()
End
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Time
End Sub
Private Sub Timer2_Timer()
If Text1.Text < Time Then
Beep: Beep: Beep
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

