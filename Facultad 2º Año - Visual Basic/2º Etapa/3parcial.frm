VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "OK"
      Height          =   495
      Left            =   6120
      TabIndex        =   16
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "3parcial.frx":0000
      Left            =   360
      List            =   "3parcial.frx":0002
      TabIndex        =   13
      Top             =   3120
      Width           =   7695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remover Item "
      Height          =   495
      Left            =   6480
      TabIndex        =   10
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Agregar Item "
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "3parcial.frx":0004
      Left            =   2520
      List            =   "3parcial.frx":0017
      TabIndex        =   4
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Total de Libros prestados: "
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   4920
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Cod.:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Apellido y Nombre"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Nro. Socio"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sum As Integer
Private Sub Combo1_Click()
Text1.Text = Combo1.ListIndex + 1
End Sub

Private Sub Command3_Click()
List1.AddItem Text2.Text & " - " & Text3.Text
sum = sum + 1
Label7.Caption = sum
End Sub

Private Sub Command4_Click()
List1.RemoveItem List1.ListIndex
Text2.Text = ""
Text3.Text = ""
sum = sum - 1
Label7.Caption = sum
End Sub

Private Sub Command5_Click()
List1.Clear
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
sum = 0
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()
Label1.Caption = "Prestamo Numero: 1"
sum = 0
End Sub

