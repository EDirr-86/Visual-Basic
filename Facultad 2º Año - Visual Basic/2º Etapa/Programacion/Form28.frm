VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Limpiar"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese dos cadenas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.TextBox Text3 
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1920
         Width           =   4455
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Texto a buscar"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Cadena original"
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
         Left            =   120
         TabIndex        =   4
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
Private Sub Command1_Click()
Text3 = "Subcadena encontrada en la posicion: " & InStr(Text1, Text2)
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Text1 = ""
Text2 = ""
Text3 = ""
End Sub
