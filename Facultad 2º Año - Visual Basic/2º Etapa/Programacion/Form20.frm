VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   4200
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Pago Mensual"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Asumiendo un interes de 9.5% a 10 años"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Precio de Venta"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
End
End Sub

Private Sub command1_click()
Call Pago
MsgBox (Pago)
End Sub

Public Function Pago() As Double
Pago = Pmt(Rate:=0.095, NPer:=120, PV:=-1900)
End Function
