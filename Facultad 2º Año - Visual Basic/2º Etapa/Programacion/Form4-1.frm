VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Abrir Formulario &No Modal"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abrir Formulario &Modal"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
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

Private Sub Command2_Click()
Form2.Show 1
End Sub

Private Sub Command3_Click()
Form2.Show 0
End Sub
