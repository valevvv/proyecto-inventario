VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmregistros 
   Caption         =   "Registro de Inventario"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16020
   Icon            =   "frmregistros.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmregistros.frx":1084A
   ScaleHeight     =   730
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1068
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmregistros.frx":3808D
      Height          =   6135
      Left            =   5400
      TabIndex        =   15
      Top             =   4440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10821
      _Version        =   393216
      HeadLines       =   2
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9
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
   Begin VB.CommandButton cmdnuevo 
      Height          =   735
      Left            =   1080
      Picture         =   "frmregistros.frx":380A2
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdG 
      Height          =   735
      Left            =   1080
      Picture         =   "frmregistros.frx":391DA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdM 
      Height          =   735
      Left            =   1080
      Picture         =   "frmregistros.frx":3A5CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdC 
      Height          =   735
      Left            =   1080
      Picture         =   "frmregistros.frx":3BB1C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9960
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8280
      Top             =   11400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=inventario.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=inventario.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Inventario"
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
   Begin VB.TextBox txtcant 
      DataField       =   "Cant_disponible"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtcategoria 
      DataField       =   "Categoria"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7200
      Top             =   11280
   End
   Begin VB.TextBox txtfecha 
      DataField       =   "Fecha"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtvaloracion 
      DataField       =   "Valoracion"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtid 
      DataField       =   "Id_Ropa"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Valoración total de productos en el inventario"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10080
      TabIndex        =   16
      Top             =   3240
      Width           =   5175
      Begin VB.Label lblval 
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   17
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD DISPONIBLE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4440
      TabIndex        =   9
      Top             =   3600
      Width           =   3045
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORIA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   6000
      TabIndex        =   6
      Top             =   2520
      Width           =   810
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   10200
      TabIndex        =   5
      Top             =   2640
      Width           =   4755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALORACION:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   360
      TabIndex        =   3
      Top             =   4560
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID  DE ROPA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   1980
   End
End
Attribute VB_Name = "frmregistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdC_Click()
Unload Me
End Sub
Private Sub cmdG_Click()
txtId.SetFocus
If txtId.Text <> "" Or txtvaloracion.Text <> "" Or txtcategoria.Text <> "" Or txtcant.Text <> "" Or txtFecha.Text <> "" Then
Adodc1.Recordset.Update
MsgBox "Se ha guardado correctamente"
Else
mensaje = MsgBox("Rellena las casillas", vbCritical, "Todo Jeans")
End If
End Sub

Private Sub cmdM_Click()
txtId.SetFocus
If txtId.Text <> "" Or txtvaloracion.Text <> "" Or txtcategoria.Text <> "" Or txtcant.Text <> "" Or txtFecha.Text <> "" Then
Adodc1.Recordset.Update
End If
End Sub

Private Sub cmdnuevo_Click()
txtId.SetFocus
On Error GoTo salida
Adodc1.Recordset.AddNew
MsgBox "Clic a lado del nombre para agregar", vbInformation, "Todo Jeans"
Exit Sub
salida:
MsgBox "Has dando clic dos veces en nuevo registro, tienes que agregar algo", vbCritical, "Todo Jeans"
End Sub

Private Sub Form_Load()
formatoropa
End Sub
Private Sub formatoropa()
DataGrid1.Columns(0).Width = 70
DataGrid1.Columns(1).Width = 70
DataGrid1.Columns(2).Width = 70
DataGrid1.Columns(3).Width = 80
DataGrid1.Columns(4).Width = 100
DataGrid1.Columns(5).Width = 70
End Sub
Private Sub Timer1_Timer()
Dim Cn As New ADODB.Connection
Cn.CursorLocation = adUseClient
Cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & App.Path & "\Inventario.accdb;Persist Security Info=False"
Dim rs As New ADODB.Recordset
rs.Open "select sum(valoracion)from inventario", Cn, adOpenStatic, adLockOptimistic
lblval.Caption = rs(0)
lblfecha.Caption = Date
End Sub


