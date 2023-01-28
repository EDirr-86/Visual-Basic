VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Resultado"
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   3735
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Total:"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Calcular"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Valor Porcentual"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call cal
Label1 = Text2 & " %"
End Sub

Private Sub Command2_Click()
End
End Sub

Public Sub cal()
Text3 = Val(Text1 * Text2) / 100

Text4 = Val(Text1 * Text2) / 100 + Val(Text1)

End Sub
