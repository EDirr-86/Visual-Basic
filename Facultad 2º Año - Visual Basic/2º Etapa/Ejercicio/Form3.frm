VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   5550
   ClientTop       =   2115
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   10470
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   5160
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar "
      Height          =   4575
      Left            =   4680
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1935
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   5
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4560
      TabIndex        =   15
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   435
      Left            =   8640
      TabIndex        =   9
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   435
      Left            =   240
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "ID"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Precio"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Descripcion"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim agregar As Boolean
Private Sub Command1_Click()
rs.MovePrevious
If rs.BOF Then
    rs.MoveFirst
End If
mostrar
End Sub

Private Sub Command2_Click()
agregar = True
limpiar
Text1.SetFocus
Command4.Enabled = False
Command5.Enabled = False
Command2.Enabled = False
Command6.Enabled = True
Command7.Enabled = True
rs.Open "select id from producto order by id asc", cn, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
rs.MoveLast
Text5 = rs.Fields("id") + 1
Else
Text5 = 1
End If
rs.Close
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
End Sub

Private Sub Command3_Click()
rs.MoveNext
If rs.EOF Then
    rs.MoveLast
End If
mostrar
End Sub

Public Sub mostrar()
Text1 = rs.Fields("nombre")
Text2 = rs.Fields("description")
Text3 = rs.Fields("cantidad")
Text4 = rs.Fields("precio")
Text5 = rs.Fields("id")
End Sub

Private Sub Command4_Click()
agregar = False
If Text5 <> "" Then
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Else
MsgBox ("Debe buscar un registro")
End If
Command4.Enabled = False
Command5.Enabled = False
Command2.Enabled = False
Command6.Enabled = True
Command7.Enabled = True
End Sub

Private Sub Command5_Click()
If MsgBox("Esta seguro que desea eliminar", vbYesNo, "Atencion") = vbYes Then
rs.Open "select * from producto where id = " & Val(Text5), cn, adOpenDynamic, adLockOptimistic
limpiar
rs.Delete
rs.Close
End If
End Sub

Private Sub Command6_Click()
If agregar = True Then
rs.Open "producto", cn, adOpenDynamic, adLockOptimistic
rs.AddNew
pasardatos
rs.Update
limpiar
Else
rs.Open "select * from producto where id= " & Val(Text5), cn, adOpenDynamic, adLockOptimistic
pasardatos
rs.Update
limpiar
End If
rs.Close
Command4.Enabled = True
Command5.Enabled = True
Command2.Enabled = True
Command6.Enabled = False
Command7.Enabled = False

End Sub

Private Sub Command8_Click()
End
End Sub

Private Sub Command9_Click()
Frame1.Visible = True
Text6.SetFocus
Text6 = ""
MSHFlexGrid1.Clear
Command4.Enabled = True
End Sub

Private Sub Form_Load()
rs.Open "producto", cn, adOpenStatic, adLockReadOnly
rs.Close
End Sub

Public Sub limpiar()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
End Sub

Public Sub pasardatos()
' convertir los datos segun formato de la base de datos acces
' funciones val() entero cdbl() double, cvdate() fecha
rs.Fields("id") = Val(Text5.Text)
rs.Fields("nombre") = Text1.Text
rs.Fields("description") = Text2.Text
rs.Fields("cantidad") = Val(Text3.Text)
rs.Fields("precio") = CDbl(Text4.Text)
End Sub

Private Sub MSHFlexGrid1_Click()
Text1 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
Text5 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
Text2 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
Text3 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)
Text4 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)
Frame1.Visible = False
End Sub

Private Sub Text6_Change()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 1
If Text6 <> "" Then
rs.Open "select * from producto where id = " & Val(Text6), cn, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
    MSHFlexGrid1.AddItem rs.Fields(0) & vbTab & rs.Fields(1) & vbTab & rs.Fields(2) & vbTab & rs.Fields(3) & vbTab & rs.Fields(4)
Else: MsgBox ("No existe el registro")
End If
rs.Close
End If
End Sub
