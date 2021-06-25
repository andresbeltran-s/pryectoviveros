VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form modificar 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   Picture         =   "modificar.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   735
      Left            =   9600
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label foto1 
         Caption         =   "Label3"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "FOTO"
      Height          =   2775
      Left            =   8280
      TabIndex        =   13
      Top             =   480
      Width           =   3015
      Begin VB.Image foto 
         Height          =   2055
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Explorar imagen"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdnuevo2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdguar 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   6
      Top             =   3360
      Width           =   2295
      Begin VB.TextBox canti1 
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
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
      Begin VB.TextBox precio1 
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
      Begin VB.TextBox descri1 
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
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   3975
      Begin VB.TextBox nom1 
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
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog abrir 
      Left            =   4680
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "AGREGAR / MODIFICAR"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   11
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7800
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "modificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdguar_Click()
plan
With rsp
.Find "PLANTAID='" & Label1.Caption & "'"
!ESPECIE = nom1.Text
!DESCRIPCION = descri1.Text
!PRECIO = precio1.Text
!stock = canti1.Text
.Update
End With
End Sub

Private Sub cmdnuevo2_Click()
plan
With rsp
.AddNew
!ESPECIE = nom1.Text
!DESCRIPCION = descri1.Text
!PRECIO = precio1.Text
!stock = canti1.Text
!foto = foto1.Caption
.Update
End With


FileCopy abrir.FileName, App.Path & "\" & abrir.FileTitle
MsgBox "Ya se agrego a la base de datos"
End Sub

Private Sub Command8_Click()
abrir.ShowOpen
foto.Picture = LoadPicture(abrir.FileName)
foto1.Caption = abrir.FileTitle
If foto1.Caption = "" Then
MsgBox "Selecione una imagen"
Else
foto1.Caption = abrir.FileTitle
End If
End Sub

Private Sub Form_Load()
cmdguar.Picture = LoadPicture(App.Path & "\img\guardarcam.gif")
cmdnuevo2.Picture = LoadPicture(App.Path & "\img\nuevo2.gif")
Frame1.BackColor = RGB(220, 220, 220)
Frame2.BackColor = RGB(220, 220, 220)
Frame3.BackColor = RGB(220, 220, 220)
Frame4.BackColor = RGB(220, 220, 220)
End Sub

