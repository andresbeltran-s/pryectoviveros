Attribute VB_Name = "declaraciones"
Option Explicit
Global BASE As New ADODB.Connection
Global rsadmin As New ADODB.Recordset
Global rsp As New ADODB.Recordset
Global detallefac As New ADODB.Recordset
Global rsdetalles As New ADODB.Recordset
Global rspedido As New ADODB.Recordset
Global rsdetalleel As New ADODB.Recordset
Global rspedidoel As New ADODB.Recordset
Public tot As Double
Public fila As Integer
Public a As Double
Global rsplanta2 As New ADODB.Recordset

