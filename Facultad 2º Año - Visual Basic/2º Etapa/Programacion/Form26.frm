VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpiar"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Factorial"
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Numero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Calcular"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Fac
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
Text1 = ""
Text1.SetFocus
End Sub
Public Sub Fac()
aux = 1
For cal = 1 To Val(Text1) Step 1
aux = cal * aux
Next
MsgBox ("El factorial de " & Val(Text1) & " es: " & aux)
End Sub
