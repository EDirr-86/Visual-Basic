Attribute VB_Name = "Module1"
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Sub main()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\emi\Desktop\Ejercicio\Sistema StockII.mdb;Persist Security Info=False"
cn.Open
Form1.Show
End Sub
