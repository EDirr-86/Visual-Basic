VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   6705
   ClientTop       =   1185
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   7725
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   5880
      TabIndex        =   26
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Resultado"
      Height          =   495
      Left            =   3960
      TabIndex        =   25
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   2040
      TabIndex        =   24
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Agregar"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6240
      TabIndex        =   22
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Forma de Pago"
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   6600
      Width           =   3615
      Begin VB.OptionButton Option2 
         Caption         =   "Tarjeta"
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Contado"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   7455
      Begin VB.ListBox List4 
         Height          =   2595
         ItemData        =   "Form1Prac.frx":0000
         Left            =   6120
         List            =   "Form1Prac.frx":0002
         TabIndex        =   27
         Top             =   960
         Width           =   1095
      End
      Begin VB.ListBox List3 
         Height          =   2595
         ItemData        =   "Form1Prac.frx":0004
         Left            =   4920
         List            =   "Form1Prac.frx":0006
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   2595
         ItemData        =   "Form1Prac.frx":0008
         Left            =   3360
         List            =   "Form1Prac.frx":000A
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   2595
         ItemData        =   "Form1Prac.frx":000C
         Left            =   240
         List            =   "Form1Prac.frx":000E
         TabIndex        =   9
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4920
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1Prac.frx":0010
         Left            =   240
         List            =   "Form1Prac.frx":0029
         TabIndex        =   28
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Sub-Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Unidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Precio Unit."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7455
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Telefono:"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido, Nombre:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Total Final:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Double
Dim remover As Double
Private Sub Command1_Click()
List1.AddItem Combo1.Text
List2.AddItem Text3
List3.AddItem Text4
List4.AddItem Val(Text3) * CDbl(Text4)
total = Val(Text3) * CDbl(Text4)
Text5.Text = total + Val(Text5)
End Sub

Private Sub Command2_Click()
remover = Val(List4.List(List1.ListIndex))
Text5.Text = Val(Text5.Text) - remover
List4.RemoveItem List1.ListIndex
List3.RemoveItem List1.ListIndex
List2.RemoveItem List1.ListIndex
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Label3_Click()
Label3.Caption = Date
End Sub
