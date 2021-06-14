VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmControl 
   Caption         =   "CONTROL DE PRODUCTOS"
   ClientHeight    =   11145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13215
   Icon            =   "frmcontrol.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmcontrol.frx":1084A
   ScaleHeight     =   743
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   881
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      DataField       =   "Id_Proveedores"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmcontrol.frx":3009C
      Left            =   10320
      List            =   "frmcontrol.frx":3009E
      TabIndex        =   22
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Id_Ropa"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   12298
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmcontrol.frx":300A0
      Left            =   1800
      List            =   "frmcontrol.frx":300B0
      TabIndex        =   21
      Top             =   1800
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmcontrol.frx":300C8
      Height          =   6135
      Left            =   1680
      TabIndex        =   14
      Top             =   3600
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10821
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   2
      RowHeight       =   18
      RowDividerStyle =   5
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
         Name            =   "Consolas"
         Size            =   9.75
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
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   8880
      Picture         =   "frmcontrol.frx":300DD
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9960
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   5400
      Picture         =   "frmcontrol.frx":30ECA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   6840
      Picture         =   "frmcontrol.frx":31CA2
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   3360
      Picture         =   "frmcontrol.frx":32A65
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9960
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   495
      Left            =   10320
      Picture         =   "frmcontrol.frx":33899
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "MOSTRAR DATOS"
      Height          =   735
      Left            =   2640
      TabIndex        =   15
      Top             =   11880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtMarca 
      DataField       =   "Talla"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtDescripcion 
      DataField       =   "Descripcion"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtPrecio 
      DataField       =   "Precio"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtCategoria 
      DataField       =   "Categoria"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   1680
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9480
      Top             =   11640
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "Ropa"
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
   Begin VB.CommandButton cmdModificar 
      Height          =   495
      Left            =   5880
      Picture         =   "frmcontrol.frx":3456B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdeliminar 
      Height          =   495
      Left            =   8160
      Picture         =   "frmcontrol.frx":354A4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdguardar 
      Height          =   495
      Left            =   3720
      Picture         =   "frmcontrol.frx":362EA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdnuevo 
      Height          =   495
      Left            =   1800
      Picture         =   "frmcontrol.frx":3712B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID PROVEEDORES:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8160
      TabIndex        =   9
      Top             =   2520
      Width           =   1830
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TALLA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4440
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8160
      TabIndex        =   3
      Top             =   1800
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORIA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID ROPA:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   945
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub DesabilitarControles()
    txtcategoria.Enabled = False
    txtPrecio.Enabled = False
    txtDescripcion.Enabled = False
    txtMarca.Enabled = False
   Combo2.Enabled = False
    cmdguardar.Enabled = False
    Combo1.Enabled = False
    cmdnuevo.Enabled = False
    
End Sub
Sub HabilitarControles()
    txtcategoria.Enabled = True
    txtPrecio.Enabled = True
    txtDescripcion.Enabled = True
    txtMarca.Enabled = True
   Combo2.Enabled = True
 
End Sub
Private Sub cmCancelar_Click()
Adodc1.Recordset.CancelUpdate
Call DesabilitarControles
cmdeliminar.Enabled = False
End Sub
Private Sub cmdeliminar_Click()
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
   Adodc1.Recordset.MoveLast
 End If
End Sub

Private Sub cmdguardar_Click()
If txtcategoria.Text <> "" Or txtPrecio.Text <> "" Or txtDescripcion.Text <> "" Or txtMarca.Text <> "" Or Combo2.Text <> "" Then
Adodc1.Recordset.Update
MsgBox "Se ha guardado correctamente"
Else
mensaje = MsgBox("Rellena las casillas", vbCritical, "Todo Jeans")
End If
End Sub

Private Sub cmdModificar_Click()
Call HabilitarControles
cmdnuevo.Enabled = True
cmdguardar.Enabled = True
txtcategoria.SetFocus
If txtcategoria.Text <> "" Or txtPrecio.Text <> "" Or txtDescripcion.Text <> "" Or txtMarca.Text <> "" Or Combo2.Text <> "" Then
Adodc1.Recordset.Update
Else
mensaje = MsgBox("Rellena las casillas", vbCritical, "Todo Jeans")
End If
End Sub

Private Sub cmdMostrar_Click()
'Dim cn As New ADODB.Connection 'Creamos el objeto Connection.
    'Dim rs As New ADODB.Recordset 'Creamos el objeto Recordset.
    
    'Abrimos la base de datos "agenda.mdb".
    'cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\inventario.accdb" 'Recuerden poner el lugar donde esta el archivo de datos en su pc
    'rs.Source = "Ropa" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
    'rs.Open "select * from Ropa", cn
End Sub
Private Sub cmdnuevo_Click()
Combo1.Enabled = True
txtcategoria.SetFocus
On Error GoTo salida
Adodc1.Recordset.AddNew
MsgBox "Coloque el producto", vbInformation, "Todo Jeans"
Exit Sub
salida:
MsgBox "Has dando clic dos veces en nuevo registro, tienes que agregar algo", vbCritical, "Todo Jeans"
End Sub

Private Sub cmdSalir_Click()
If MsgBox("¿Esta seguro de cerrar?", vbQuestion + vbYesNo) = vbYes Then
Unload Me
End If
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveLast
End Sub


Private Sub Form_Load()
tablaproveedores
Call DesabilitarControles
formatoropa
While rsproveedores.EOF = False
    Combo2.AddItem (rsproveedores!Id_proveedores)
    rsproveedores.MoveNext
Wend
End Sub
Private Sub formatoropa()
DataGrid1.Columns(0).Width = 70
DataGrid1.Columns(1).Width = 150
DataGrid1.Columns(2).Width = 50
DataGrid1.Columns(3).Width = 200
DataGrid1.Columns(4).Width = 50
DataGrid1.Columns(5).Width = 105
End Sub

