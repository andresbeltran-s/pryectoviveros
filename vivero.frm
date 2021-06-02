VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000010&
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16830
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   16830
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   615
      Left            =   9840
      Top             =   8160
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Usuario\Desktop\proyectoviveros\basevivero.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Usuario\Desktop\proyectoviveros\basevivero.mdb;Persist Security Info=False"
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
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   11640
      TabIndex        =   16
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Frame Frame6 
      Height          =   855
      Left            =   1920
      TabIndex        =   11
      Top             =   8040
      Width           =   7455
      Begin VB.CommandButton Command4 
         Caption         =   "Iniciar sesión"
         Height          =   495
         Left            =   5400
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Insertar"
         Height          =   495
         Left            =   3840
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   495
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Editar"
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Drescripción"
      Height          =   855
      Left            =   7440
      TabIndex        =   9
      Top             =   2160
      Width           =   3255
      Begin VB.TextBox txtdescri 
         Height          =   405
         Left            =   120
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Buscar "
      Height          =   855
      Left            =   7440
      TabIndex        =   8
      Top             =   720
      Width           =   3735
      Begin VB.CommandButton cmdbus 
         Caption         =   "Buscar "
         Height          =   435
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtbus 
         Height          =   405
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Total"
      Height          =   855
      Left            =   5760
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
      Begin VB.TextBox txttot 
         BackColor       =   &H80000006&
         Height          =   405
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   600
      TabIndex        =   5
      Top             =   5760
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   -2147483632
      ForeColor       =   4194304
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cantidad"
      Height          =   1215
      Left            =   720
      TabIndex        =   3
      Top             =   4200
      Width           =   2655
      Begin VB.TextBox txtcantidad 
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Precio"
      Height          =   855
      Left            =   3480
      TabIndex        =   0
      Top             =   4200
      Width           =   2175
      Begin VB.TextBox txtprecio 
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "$"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Image Image1 
      Height          =   3720
      Left            =   600
      Picture         =   "vivero.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   6600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim cn As New ADODB.Connection
 Dim x As String
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
   


Private Sub cmdbus_Click()
plan
x = txtbus.Text
With rsp
.Find "ESPECIE='" & x & "'"
If .EOF Or .BOF Then Exit Sub
txtdescri.Text = !DESCRIPCIÓN
txtprecio.Text = !PRECIO
End With

End Sub

Private Sub Command3_Click()
 Do Until rs.EOF
        List1.AddItem rs.Fields("PRECIO") & rs.Fields("DESCRPCION") & rs.Fields("STOCK") & rs.Fields("ESPECIE")
        rs.MoveNext
    Loop
End Sub

Private Sub txtcantidad_Change()
plan
txttot.Text = Val(txtcantidad.Text) * Val(txtprecio.Text)
End Sub
