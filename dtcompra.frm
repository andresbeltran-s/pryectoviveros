VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form dtcompra 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   13050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtcantidad 
      Height          =   375
      Left            =   5160
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtprecio 
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtnom1 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Lista 
      Height          =   3735
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   20
      Cols            =   5
   End
   Begin VB.CommandButton Command1 
      Caption         =   "aceptar"
      Height          =   615
      Left            =   5760
      TabIndex        =   3
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Precio"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Cantidad"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "nombre del producto:"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "total"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "iva"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "sub total"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "dtcompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Lista.ColWidth(0) = 10
Lista.ColWidth(1) = 3000
Lista.ColAlignment(1) = 5
Lista.Col = 1
Lista.Row = 0
Lista.Text = "Producto"
Lista.ColWidth(2) = 3000
Lista.ColAlignment(2) = 5
Lista.Col = 2
Lista.Row = 0
Lista.Text = "Precio"
Lista.ColWidth(3) = 3000
Lista.ColAlignment(3) = 5
Lista.Col = 3
Lista.Row = 0
Lista.Text = "Cantidad"
Lista.ColWidth(4) = 3000
Lista.ColAlignment(4) = 5
Lista.Col = 4
Lista.Row = 0
Lista.Text = "Total Unico"
fila = 1

End Sub





Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
If KeyAscii Then
    Lista.Col = 1
    Lista.Row = fila
    Lista.Text = txtnom1.Text
    Lista.Col = 2
    Lista.Row = fila
    Lista.Text = txtprecio.Text
    Lista.Col = 3
    Lista.Row = fila
    Lista.Text = txtcantidad.Text
    a = Val(txtprecio.Text) * Val(txtcantidad.Text)
    Lista.Col = 4
    Lista.Row = fila
    Lista.Text = a
    fila = fila + 1
    txtnom1.Text = ""
    txtprecio.Text = ""
    txtcantidad.Text = ""
    txtnom1.SetFocus
End If
End Sub

Private Sub txtnom1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtprecio.SetFocus
End If
End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtcantidad.SetFocus
End If
End Sub
