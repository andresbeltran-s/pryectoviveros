VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form stock 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   735
      Left            =   8880
      TabIndex        =   4
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   795
      Left            =   6840
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtstock 
      Height          =   525
      Left            =   7920
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   8916
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
            ColumnWidth     =   104,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   104,882
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3015
   End
End
Attribute VB_Name = "stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e As String

Private Sub Command1_Click()
    With rsp
        .Find "ESPECIE='" & Text1.Text & "'"
        !stock = !stock + 1
        .UpdateBatch
        txtstock.Text = !stock
    End With
End Sub

Private Sub Command2_Click()
    With rsp
        .Find "ESPECIE='" & Text1.Text & "'"
        If !stock = 0 Then Exit Sub
        !stock = !stock - 1
        .UpdateBatch
        txtstock.Text = !stock
    End With
End Sub

Private Sub DataGrid1_Click()
    Text1.Text = DataGrid1.Columns(1).Text
    With rsp
        .Find "ESPECIE='" & DataGrid1.Columns(1).Text & "'"
        e = !foto
    End With
    Image1.Picture = LoadPicture(App.Path & "/plantas/" & e)
    txtstock.Text = DataGrid1.Columns(4).Text
End Sub

Private Sub Form_Load()
With rsp
If .State = 1 Then .Close
    .Open "Select * From PLANTA Where [STOCK] <= 1 "
    Set DataGrid1.DataSource = rsp
End With
End Sub


