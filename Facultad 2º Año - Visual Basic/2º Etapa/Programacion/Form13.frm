VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   240
      Max             =   255
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Consulte aqui su duda       ---------------------------------->       ---------------------------------->"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "ASCII "
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
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Alt +"
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
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text3.Text = HScroll1.Min
Text1.Text = Chr(HScroll1.Min)
End Sub
Private Sub HScroll1_Change()
Text3.Text = HScroll1.Value
Text1.Text = Chr(HScroll1.Value)
End Sub
Private Sub Text1_Change()
If Text1.Text <> "" Then
Text3.Text = Asc(Text1.Text)
HScroll1.Value = Text3.Text
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If MsgBox(KeyAscii, vbApplicationModal, "Character : ") = vbOK Then
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.Text = Chr(Text3.Text)
End If
End Sub

