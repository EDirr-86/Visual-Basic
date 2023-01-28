VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Left            =   5280
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Accion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   3375
      Begin VB.CommandButton Command3 
         Caption         =   "&Limpiar"
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
         Left            =   960
         TabIndex        =   8
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Eliminar"
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
         Left            =   1680
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
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
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
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
         Height          =   405
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre"
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
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "Form12.frx":0000
      Left            =   240
      List            =   "Form12.frx":0002
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Lista"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Combo"
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
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Combo1.AddItem Text1.Text
List1.AddItem Text1.Text

Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click()
List1.Clear
Combo1.Clear
End Sub

Private Sub Command4_Click()
End
End Sub
Private Sub Command2_Click()

If Combo1.ListIndex <> -1 Then
    With Combo1
    List1.RemoveItem .ListIndex
    .RemoveItem .ListIndex
    End With
ElseIf List1.ListIndex <> -1 Then
    With List1
    Combo1.RemoveItem .ListIndex
    .RemoveItem .ListIndex
    End With

Else
MsgBox "No se ha seleccionando nada"
End If

End Sub


