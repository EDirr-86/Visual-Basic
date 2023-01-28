VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calculo del IVA"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Total"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Sub-Total"
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
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "+ 21%"
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
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   """Se confirma con ENTER"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1 = ""
Text2 = ""
Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 > 0 Then
        Text2 = Val((Text1 * 21) / 100 + Text1)
    End If
End If
End Sub
