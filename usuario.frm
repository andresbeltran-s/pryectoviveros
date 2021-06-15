VERSION 5.00
Begin VB.Form usuario 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   Picture         =   "usuario.frx":0000
   ScaleHeight     =   4785
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdeli 
      Height          =   735
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmd12 
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ComboBox combo1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "usuario.frx":DF25
      Left            =   3120
      List            =   "usuario.frx":DF2F
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtconnue 
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox txtadnue 
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
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Permiso"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd12_Click()
admin
With rsadmin
.AddNew
!NOMBRE = txtadnue.Text
!CONTRASEÑA = txtconnue.Text
!PERMISO = combo1.Text
.UpdateBatch
End With
End Sub

Private Sub cmdeli_Click()
txtadnue.Text = ""
 txtconnue.Text = ""
 combo1.Text = "seleccionar"
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
cmd12.Picture = LoadPicture(App.Path & "\img\agregra1.gif")
cmdeli.Picture = LoadPicture(App.Path & "\img\limpiar.gif")
Command1.Picture = LoadPicture(App.Path & "\img\cerrar.gif")
End Sub
