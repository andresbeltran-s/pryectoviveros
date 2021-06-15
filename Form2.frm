VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form registro 
   Caption         =   "Registro"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17820
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   7875
   ScaleWidth      =   17820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "&Detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   6975
      Begin VB.TextBox Text5 
         DataField       =   "STOCK"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         DataField       =   "PLANTAID"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         DataField       =   "ESPECIE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         DataField       =   "DESCRIPCION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         DataField       =   "PRECIO"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton Command7 
         Caption         =   ">"
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
         Left            =   5520
         TabIndex        =   10
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<"
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
         Left            =   4440
         TabIndex        =   9
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   ">|"
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
         Left            =   6240
         TabIndex        =   8
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "|<"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label lblname 
         Caption         =   "........"
         DataField       =   "FOTO"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FOTO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLANTAID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESPECIE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCIÓN:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   915
      End
   End
   Begin VB.Frame frm1 
      Caption         =   "FOTO"
      Height          =   3615
      Left            =   11160
      TabIndex        =   5
      Top             =   600
      Width           =   3375
      Begin VB.Image foto 
         Height          =   2655
         Left            =   360
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdcerrarse 
      Caption         =   "Cerrar Sesión"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   2445
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Eliminar Planta"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar Planta"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modificar Planta"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   600
      Top             =   6120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
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
      Caption         =   "Lista de Plantas"
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
   Begin MSComDlg.CommonDialog abrir 
      Left            =   3240
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LISTA DE PLANTAS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   8640
      TabIndex        =   4
      Top             =   6840
      Width           =   5295
   End
End
Attribute VB_Name = "registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As String
Private Sub cmdcerrarse_Click()
Me.Hide
frmacceso.Show
principal.Label4 = "false"
End Sub

Private Sub Command1_Click()
modificar.Show
With Adodc1.Recordset
modificar.nom1.Text = !ESPECIE
modificar.canti1.Text = !stock
modificar.precio1.Text = !PRECIO
modificar.descri1.Text = !DESCRIPCION
modificar.Label1.Caption = !PLANTAID
End With
modificar.cmdnuevo2.Enabled = False
With modificar
.cmdnuevo2.Visible = False
.Command8.Visible = False
End With
End Sub

Private Sub Command2_Click()
modificar.Show
modificar.cmdguar.Enabled = False
With modificar
.cmdguar.Visible = False
End With
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveFirst

foto.Picture = LoadPicture(x & "\" & lblname.Caption)
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveLast

foto.Picture = LoadPicture(x & "\" & lblname.Caption)
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If

foto.Picture = LoadPicture(x & "\" & lblname.Caption)
End Sub


Private Sub Command7_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveLast
End If
If lblname.Caption = "" Then
    MsgBox "no imagen"
    Else
foto.Picture = LoadPicture(x & "\" & lblname.Caption)
End If
End Sub

Private Sub Form_Load()
Command1.Picture = LoadPicture(App.Path & "\img\editar3.gif")
Command2.Picture = LoadPicture(App.Path & "\img\agregar.gif")
Command3.Picture = LoadPicture(App.Path & "\img\eli2.gif")
cmdcerrarse.Picture = LoadPicture(App.Path & "\img\cerrarsesion.gif")
plan

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\basevivero.mdb;Persist Security Info=False"

Adodc1.CursorType = adOpenDynamic
Adodc1.RecordSource = "select * from PLANTA"
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
x = App.Path & "\plantas"
foto.Picture = LoadPicture(x & "\" & lblname.Caption)
End Sub

Sub a()
modificar.Show
plan
With rsp
modificar.nom1.Text = !ESPECIE
modificar.canti1.Text = !stock
modificar.precio1.Text = !PRECIO
modificar.descri1.Text = !DESCRIPCIÓN
modificar.Label1.Caption = !PLANTAID
End With
End Sub


