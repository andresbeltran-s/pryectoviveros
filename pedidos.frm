VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form pedidos 
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox Hasta 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d/M/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Desde 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d/M/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   3
      EndProperty
      Height          =   405
      Left            =   2640
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "pedidos.frx":0000
      Left            =   360
      List            =   "pedidos.frx":000D
      TabIndex        =   4
      Text            =   "Opciones"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   4320
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2566
      _Version        =   393216
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3413
      _Version        =   393216
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
   Begin VB.Label cod 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   10800
      TabIndex        =   2
      Top             =   1560
      Width           =   45
   End
End
Attribute VB_Name = "pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Integer
Private Sub cmdeliminar_Click()
    With rspedidoel
        .AddNew
        !CODIGO_PEDIDO = DataGrid1.Columns(0).Text
        !HORA = DataGrid1.Columns(1).Text
        !FECHA = DataGrid1.Columns(2).Text
        !C_EMPLEADOS = DataGrid1.Columns(3).Text
        !Total = DataGrid1.Columns(4).Text
        !TOTAL_IVA = CDbl(DataGrid1.Columns(5).Text)
        .UpdateBatch
    End With
    q = rsdetalles.RecordCount
    For x = 1 To q
    With rsdetalleel
        .Requery
        .AddNew
        !DETALLE_ID = DataGrid2.Columns(0).Text
        !CODIGO_PEDIDO = DataGrid2.Columns(1).Text
        !PLANTA_ID = DataGrid2.Columns(2).Text
        !PLANTA = DataGrid2.Columns(3).Text
        !CANTIDAD = DataGrid2.Columns(4).Text
        !TOTALD = DataGrid2.Columns(5).Text
        .UpdateBatch
    End With
    rsdetalles.MoveNext
    Next
    With rspedido
        .Delete
        .MoveFirst
    End With
    For x = 1 To q
    With rsdetalles
        .Requery
        .Delete
        .MoveNext
        If .EOF Then Exit Sub
    End With
    Next
End Sub





Private Sub Combo1_Click()
    If combo1.Text = "Desde" Then
        Desde.Visible = True
        Hasta.Visible = False
    End If
    If combo1.Text = "Hasta" Then
        Desde.Visible = False
        Hasta.Visible = True
    End If
    If combo1.Text = "Desde/Hasta" Then
        Desde.Visible = True
        Hasta.Visible = True
    End If
End Sub

Private Sub Command1_Click()
    Dim s, a As Date
    With rspedido
        If combo1.Text = "Desde" Then
            If .State = 1 Then .Close
            s = "#" & Desde.Text & "#"
            .Open "Select * From PEDIDO Where ((PEDIDO.[FECHA])> " & s & ")"
            Set DataGrid1.DataSource = rspedido
        End If
        If combo1.Text = "Hasta" Then
            If .State = 1 Then .Close
            a = "#" & Hasta.Text & "#"
            .Open "Select * From PEDIDO Where ((PEDIDO.[FECHA])< " & a & ")"
            Set DataGrid1.DataSource = rspedido
        End If
        If combo1.Text = "Desde/Hasta" Then
            If .State = 1 Then .Close
            s = "#" & Desde.Text & "#"
            a = "#" & Hasta.Text & "#"
            .Open "Select * From PEDIDO Where ((PEDIDO.[FECHA])>= " & s & ") AND ((PEDIDO.[FECHA])<= " & a & ")"
            Set DataGrid1.DataSource = rspedido
        End If
    End With
End Sub

Private Sub DataGrid1_Click()
    cod = DataGrid1.Columns(0).Text
    With rsdetalles
        Dim s As String
        s = "%" & cod.Caption & "%"
        If .State = 1 Then .Close
        .Open "Select * From DETALLE_PEDIDO Where [CODIGO_PEDIDO]Like '" & s & "'"
        Set DataGrid2.DataSource = rsdetalles
    End With
End Sub

Private Sub Form_Load()
pedido
detalles
pedidoel
detalleel
Set DataGrid1.DataSource = rspedido

End Sub
