VERSION 5.00
Begin VB.Form frmacceso 
   BackColor       =   &H80000011&
   Caption         =   "Form2"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475
   LinkTopic       =   "Form2"
   Picture         =   "frmacceso.frx":0000
   ScaleHeight     =   4965
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Label1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton salir 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton limpiar 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton entrar 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   2325
   End
   Begin VB.TextBox txtcontrasena 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txtusuario 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "INICIAR SESIÓN"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   4335
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
    With rsadmin
    .Requery
    .Find "NOMBRE ='" & (txtusuario.Text) & "'"
    If .EOF Then
    MsgBox "el usuario no existe!!", vbCritical, "Error de usuario"
    txtusuario = ""
    txtusuario.SetFocus
    Else
    If !CONTRASEÑA = Trim(txtcontrasena.Text) Then
    principal.Label5.Caption = "false"
    If !PERMISO = "NO" Then registro.Command1.Enabled = False: registro.Command2.Enabled = False: registro.Command3.Enabled = False: principal.Label5.Caption = "true"
    frmacceso.Hide
    principal.Label4.Caption = "true"
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
 salir.Picture = LoadPicture(App.Path & "\img\salir.gif")
 limpiar.Picture = LoadPicture(App.Path & "\img\limpiar.gif")
 Label1.Picture = LoadPicture(App.Path & "\img\usuarioplus.gif")
 entrar.Picture = LoadPicture(App.Path & "\img\entrar.gif")
admin
End Sub

Private Sub Label1_Click()
If txtusuario.Text = "" Then MsgBox "ingrese el nombre de ususario", vbInformation, "Informacion inconclusa": txtusuario.SetFocus: Exit Sub
If txtcontrasena.Text = "" Then MsgBox "ingrese una contraseña", vbInformation, "Informacion inconclusa": txtcontrasena.SetFocus: Exit Sub
    With rsadmin
    .Requery
    .Find "NOMBRE ='" & (txtusuario.Text) & "'"
    If .EOF Then
    MsgBox "el usuario no existe!!", vbCritical, "Error de usuario"
    txtusuario = ""
    txtusuario.SetFocus
    Else
    If !CONTRASEÑA = Trim(txtcontrasena.Text) Then
    If !PERMISO <> "NO" Then usuario.Show
    
    Else
        MsgBox "clave incorrecta", vbCritical, "Error de clave"
        txtcontrasena = ""
        txtcontrasena.SetFocus
    End If
    End If
    End With

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

