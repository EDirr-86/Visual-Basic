VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   2925
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
MsgBox "Se a ejecutado el evento de inicio Initialize"
End Sub

Private Sub Form_Load()
MsgBox "Se a ejecutado el evento de inicio Load"
End Sub

Private Sub Form_Paint()
MsgBox "Se a ejecutado el evento de inicio Paint"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox "Se a ejecutado el evento de cierre QueryUnload"
End Sub

Private Sub Form_Resize()
MsgBox "Se a ejecutado el evento de inicio Resize"
End Sub

Private Sub Form_Terminate()
MsgBox "Se a ejecutado el evento de cierre Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Se a ejecutado el evento de cierre Unload"
End Sub

