VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form principal 
   BackColor       =   &H00000000&
   Caption         =   "Principal"
   ClientHeight    =   9510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   Picture         =   "vivero.frx":0000
   ScaleHeight     =   634
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "finalizar compra"
      Height          =   615
      Left            =   12240
      TabIndex        =   27
      Top             =   8760
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
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
      Left            =   12360
      Picture         =   "vivero.frx":11DE6
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7920
      Width           =   2175
   End
   Begin VB.CommandButton cmdinicio 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H8000000E&
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   9
      Top             =   4920
      Width           =   1335
      Begin VB.Label txtstok 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid tabla 
      Height          =   3255
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   5741
      _Version        =   393216
      BackColor       =   14737632
      ForeColor       =   0
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
   Begin VB.Frame Frame7 
      BackColor       =   &H8000000E&
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   7
      Top             =   7080
      Width           =   1095
      Begin VB.Label txtid 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   735
      Left            =   3240
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
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
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000E&
      Caption         =   "Drescripción"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   4
      Top             =   6120
      Width           =   4215
      Begin VB.Label txtdescri 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Buscar "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   3
      Top             =   4920
      Width           =   4215
      Begin VB.CommandButton cmdbus 
         Height          =   555
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbus 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   2
      Top             =   6120
      Width           =   1335
      Begin VB.Label txttot 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   1
      Top             =   4920
      Width           =   2175
      Begin VB.TextBox txtcantidad 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5280
      TabIndex        =   0
      Top             =   6120
      Width           =   2175
      Begin VB.Label txtprecio 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PEDIDOS"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   36
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   975
      Left            =   5880
      TabIndex        =   21
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   615
      Left            =   15480
      TabIndex        =   20
      Top             =   9480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   615
      Left            =   15480
      TabIndex        =   18
      Top             =   8520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   15480
      TabIndex        =   17
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Label2"
      Height          =   495
      Left            =   15480
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   15480
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim cn As New ADODB.Connection
 Dim x As String
 Dim z As Integer
 Dim Y As Integer
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
   

Private Sub cmdbus_Click()
plan
x = txtbus.Text
With rsp
.Find "ESPECIE='" & x & "'"
If .EOF Or .BOF Then MsgBox "Esta planta no se ecnuentra disponible": txtbus.Text = "": Exit Sub
txtdescri.Caption = !DESCRIPCION
txtprecio.Caption = !PRECIO
txtid.Caption = !PLANTAID
txtstok.Caption = !stock
End With

End Sub

Private Sub cmdinicio_Click()
frmacceso.Show
End Sub

Private Sub Command1_Click()
If txtid.Caption = "" Then Exit Sub
If Val(txtcantidad.Text) > Val(txtstok.Caption) Then MsgBox "No hay suficicente stock": Exit Sub
detallefactura
With detallefac
.AddNew
!FECHA = Date
!PLANTA_ID = txtid.Caption
!CANTIDAD = txtcantidad.Text
!TOTALD = txttot.Caption
.Update
End With
With rsp
.Find "PLANTAID='" & txtid.Caption & "'"
!stock = Val(!stock) - Val(txtcantidad.Text)
.Update
End With
txtstok.Caption = Val(txtstok.Caption) - Val(txtcantidad.Text)
detallefactura
Set tabla.DataSource = detallefac
tabla.Columns(0).Width = 180
tabla.Columns(1).Width = 180
tabla.Columns(2).Width = 180
tabla.Columns(3).Width = 180
tabla.Columns(4).Width = 180
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
detallefactura
If detallefac.EOF Or detallefac.BOF Then Command2.Enabled = False: Exit Sub

With detallefac


.MoveLast
If .EOF Or .BOF Then Command2.Enabled = False: Exit Sub
Label2.Caption = !PLANTA_ID
Label3.Caption = !CANTIDAD
.Delete
End With
Adodc1.CursorLocation = adUseClient
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\basevivero.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from PLANTA where [PLANTAID]LIKE '" & Label2.Caption & "'"

rsp.Find "PLANTAID= '" & Label2.Caption & "'"
rsp!stock = Val(rsp!stock) + Val(Label3.Caption)
rsp.Update
txtstok.Caption = Val(txtstok.Caption) + Val(txtcantidad.Text)
detallefactura
Set tabla.DataSource = detallefac
tabla.Columns(0).Width = 180
tabla.Columns(1).Width = 180
tabla.Columns(2).Width = 180
tabla.Columns(3).Width = 180
tabla.Columns(4).Width = 180
End Sub



Private Sub Command3_Click()
dtcompra.Show
End Sub

Private Sub Command4_Click()
detallefactura
Set DataReport2.DataSource = detallefac
DataReport2.Show
Set tabla.DataSource = detallefac
tabla.Columns(0).Width = 180
tabla.Columns(1).Width = 180
tabla.Columns(2).Width = 180
tabla.Columns(3).Width = 180
tabla.Columns(4).Width = 180
End Sub

Private Sub Command5_Click()
If Label4.Caption = "Label4" Then MsgBox "Inicie sesión": Exit Sub

If Label4.Caption = "true" Then
If Label5.Caption = "true" Then registro.Command1.Enabled = False: registro.Command2.Enabled = False: registro.Command3.Enabled = False
registro.Show
End If
End Sub

Private Sub Form_Load()
Frame1.BackColor = RGB(46, 139, 87)
Frame2.BackColor = RGB(46, 139, 87)
Frame3.BackColor = RGB(46, 139, 87)
Frame4.BackColor = RGB(46, 139, 87)
Frame5.BackColor = RGB(46, 139, 87)
Frame7.BackColor = RGB(46, 139, 87)
Frame8.BackColor = RGB(46, 139, 87)
Command1.Picture = LoadPicture(App.Path & "\img\anadir1.gif")
cmdbus.Picture = LoadPicture(App.Path & "\img\buscar.gif")
Command2.Picture = LoadPicture(App.Path & "\img\eliminar.gif")
Command5.Picture = LoadPicture(App.Path & "\img\listplan.gif")
Command4.Picture = LoadPicture(App.Path & "\img\reporte.gif")
cmdinicio.Picture = LoadPicture(App.Path & "\img\iniciose.gif")
detallefactura
plan
Set tabla.DataSource = detallefac
tabla.Columns(0).Width = 180
tabla.Columns(1).Width = 180
tabla.Columns(2).Width = 180
tabla.Columns(3).Width = 180
tabla.Columns(4).Width = 180
End Sub





Private Sub txtcantidad_Change()
plan
With rsp
x = txtid.Caption
.Find "PLANTAID= '" & x & "'"
If Val(txtcantidad.Text) > Val(!stock) Then MsgBox " supera la cantidad del stock": Exit Sub
txttot.Caption = Val(txtcantidad.Text) * Val(txtprecio.Caption)
End With
End Sub
