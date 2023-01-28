VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Limpiar"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      Width           =   2655
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese su nota"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Label2.Caption = ""
End Sub

Private Sub Text1_Change()
If Val(Text1) > 10 Then
Label2 = "No es una nota"
End If

If Val(Text1) = 10 Then
Label2 = "Exclente"
End If

If Val(Text1) = 9 Then
Label2 = "Muy bueno"
ElseIf Val(Text1) = 8 Then
Label2 = "Muy bueno"
End If

If Val(Text1) = 7 Then
Label2 = "Bueno"
ElseIf Val(Text1) = 6 Then
Label2 = "Bueno"
End If

If Val(Text1) = 5 Then
Label2 = "Regular"
ElseIf Val(Text1) = 4 Then
Label2 = "Regular"
End If

If Val(Text1) < 4 Then
Label2 = "Insuficiente"
End If

End Sub
