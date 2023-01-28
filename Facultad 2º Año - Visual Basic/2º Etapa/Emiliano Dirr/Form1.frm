VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Buscador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton salir2 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   5775
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   1215
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   4
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Numero Legajo"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton salir1 
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
      Left            =   240
      TabIndex        =   13
      Top             =   3480
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4800
      TabIndex        =   9
      Top             =   240
      Width           =   1935
      Begin VB.CommandButton cancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
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
         TabIndex        =   12
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton guardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
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
         TabIndex        =   11
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton agregar 
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
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2760
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1320
      List            =   "Form1.frx":0010
      TabIndex        =   7
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton buscar 
      Caption         =   "..."
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
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "dd / mm / aaa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "( Apellido, Nombres )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha de Contratación:"
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
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Cargo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Apellido y Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Legajo:"
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
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim agregarr As Boolean
Dim datecheck As Boolean
Dim firstdate As Date

Private Sub agregar_Click()
agregarr = True
agregar.Enabled = False
guardar.Enabled = True
cancelar.Enabled = True
limpiar
Text2.Enabled = True
Text2.SetFocus
Text3.Enabled = True
Combo1.Enabled = True


rs.Open "select legajo from empleados order by legajo asc", cn, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
rs.MoveLast
Text1 = rs.Fields("legajo") + 1
Else
Text1 = 1
End If
rs.Close
End Sub

Private Sub buscar_Click()
Frame2.Visible = True
Text4 = ""
grilla.Clear
buscar.Enabled = False
agregar.Enabled = False
End Sub


Private Sub cancelar_Click()
limpiar
Text2.Enabled = False
Text3.Enabled = False
Combo1.Enabled = False
agregar.Enabled = True
cancelar.Enabled = False
guardar.Enabled = False
End Sub

Private Sub grilla_Click()
Text1 = grilla.TextMatrix(grilla.Row, 0)
Text2 = grilla.TextMatrix(grilla.Row, 1)
Combo1 = grilla.TextMatrix(grilla.Row, 2)
Text3 = grilla.TextMatrix(grilla.Row, 3)
End Sub

Private Sub guardar_Click()
If Text1 <> "" And Text2 <> "" And Text3 <> "" And Combo1.ListIndex <> -1 Then
If agregarr = True Then
rs.Open "select * from empleados", cn, adOpenDynamic, adLockOptimistic
rs.AddNew
Else
rs.Open "select * from empleados where legajo" & Val(Text1), cn, adOpenDynamic, adLockOptimistic
End If
pasardatos
rs.Update
limpiar
rs.Close
guardar.Enabled = False
cancelar.Enabled = False
Else
MsgBox ("Falta Llenar campos!!!")
End If
End Sub

Private Sub salir1_Click()
End
End Sub

Private Sub salir2_Click()
Frame2.Visible = False

buscar.Enabled = True
agregar.Enabled = True
End Sub

Public Sub limpiar()
Text1 = ""
Text2 = ""
Text3 = ""
Combo1.ListIndex = -1
End Sub

Public Sub pasardatos()
rs.Fields("legajo") = Val(Text1)
rs.Fields("nombre") = Text2
rs.Fields("fecha_contra") = Text3
rs.Fields("cargo") = Combo1.List(Combo1.ListIndex)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
strValid = "qwertyuiopasdfghjklñzxcvbnm,QWERTYUIOPASDFGHJKLÑZXCVBNM " 'Los caracteres que ustedes quieran que ingresen
'KeyAscii = Asc(UCase(Chr(KeyAscii))) ---> Esto es un anexo para que todas la letras sean mayuscula =)
If KeyAscii > 26 Then 'para que no tome los botones accionales del teclado
If InStr(strValid, Chr(KeyAscii)) = 0 Then
KeyAscii = 0
End If
End If
If KeyAscii = 13 Then
KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
datecheck = True
strValid = "/1234567890" 'Los caracteres que ustedes quieran que ingresen
If KeyAscii > 26 Then 'para que no tome los botones accionales del teclado
If InStr(strValid, Chr(KeyAscii)) = 0 Then
beep
KeyAscii = 0
End If
End If
If KeyAscii = 13 Then
firstdate = CDate(Text3)
datecheck = IsDate(firstdate)
If datecheck = True Then
    If firstdate > Date Then
    Text3 = ""
    MsgBox ("La fecha a sido borrada por que es mayor a la actual")
End If
End If
End If
'If KeyAscii < Asc("/") Or KeyAscii > Asc("9") Then
'Beep
'KeyAscii = 8
'If KeyAscii <> 8 Then
'KeyAscii = 0
'End If
'End If
End Sub

Private Sub Text4_Change()
grilla.Clear
grilla.Rows = 1

If Text4 <> "" Then
rs.Open "select * from empleados where legajo =" & Val(Text4), cn, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
grilla.AddItem rs.Fields(0) & vbTab & rs.Fields(1) & vbTab & rs.Fields(2) & vbTab & rs.Fields(3)
Else
MsgBox ("No existe el registro")
End If
rs.Close
End If
End Sub
