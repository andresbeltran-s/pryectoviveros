VERSION 5.00
Begin VB.Form frmacceso 
   Caption         =   "Form2"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   LinkTopic       =   "Form2"
   ScaleHeight     =   2790
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton salir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton limpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton entrar 
      Caption         =   "Entrar"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtcontrasena 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtusuario 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label contra 
      Caption         =   "Contraseña"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label usu 
      Caption         =   "Usuario"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmacceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub entrar_Click()
If txtusuario.Text = "" Then MsgBox "ingrese el nombre de ususario", vbInformation, "Informacion inconclusa": txtusuario.SetFocus: Exit Sub
If txtcontrasena.Text = "" Then MsgBox "ingrese una contraseña", vbInformation, "Informacion inconclusa": txtcontrasena.SetFocus: Exit Sub
    With rsplanta
    .Requery
    .Find "NOMBRE ='" & (txtusuario.Text) & "'"
    If .EOF Then
    MsgBox "el usuario no existe!!", vbCritical, "Error de usuario"
    txtusuario = ""
    txtusuario.SetFocus
    Else
    If !CONTRASEÑA = Trim(txtcontrasena.Text) Then
    Load frmplanta
    frmplanta.Show
    frmacceso.Hide
    Unload frmacceso
    Else
        MsgBox "clave incorrecta", vbCritical, "Error de clave"
        txtcontrasena = ""
        txtcontrasena.SetFocus
    End If
    End If
    End With
End Sub

Private Sub Form_Load()
planta
End Sub

Private Sub limpiar_Click()
txtusuario = ""
txtcontrasena = ""
txtusuario.SetFocus
End Sub

Private Sub salir_Click()
Dim mensaje As String
mensaje = MsgBox("Confirmar salir de la aplicación", vbYesNo, "Salida")
If mensaje = vbYes Then
mensaje = MsgBox("gracias, hasta pronto", vbInformation, "salida")
Unload Me
End
Else
mensaje = MsgBox("aplicación corriendo nuevamente", vbInformation, " mensaje")
End If
End Sub
