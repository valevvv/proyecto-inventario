VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEditar 
   Caption         =   "Editar Usuario"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16140
   Icon            =   "frmEditar.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmEditar.frx":1084A
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1076
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEditar.frx":2019E
      Height          =   2655
      Left            =   4200
      TabIndex        =   12
      Top             =   600
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
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
   Begin VB.CommandButton cmdC 
      Height          =   735
      Left            =   13920
      Picture         =   "frmEditar.frx":201B3
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdM 
      Height          =   735
      Left            =   11040
      Picture         =   "frmEditar.frx":215E6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdG 
      Height          =   735
      Left            =   4800
      Picture         =   "frmEditar.frx":22B38
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ComboBox cmbCuenta 
      DataField       =   "TipoCuenta"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmEditar.frx":23F28
      Left            =   2160
      List            =   "frmEditar.frx":23F32
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2400
      Top             =   4920
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      RecordSource    =   "Usuario"
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
   Begin VB.CommandButton cmdnuevo 
      Height          =   735
      Left            =   7800
      Picture         =   "frmEditar.frx":23F4E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtContraseña 
      DataField       =   "Contraseña"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   2
      Top             =   2085
      Width           =   1815
   End
   Begin VB.TextBox txtCorreo 
      DataField       =   "Correo"
      DataSource      =   "Adodc1"
      Height          =   360
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtNombre 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cuenta"
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
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
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
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Correo"
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
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
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
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub DesabilitarControles()
    txtNombre.Enabled = False
    txtCorreo.Enabled = False
    txtContraseña.Enabled = False
    cmbCuenta.Enabled = False
    cmdnuevo.Enabled = False
    cmdG.Enabled = False
    cmdM.Enabled = True
End Sub
Sub HabilitarControles()
        cmbCuenta.Enabled = True
        txtNombre.Enabled = True
        txtCorreo.Enabled = True
        txtContraseña.Enabled = True
        cmdG.Enabled = True
        cmdnuevo.Enabled = True
End Sub

Private Sub cmdC_Click()
Unload Me
End Sub
Private Sub cmdG_Click()
'Call DesabilitarControles
txtNombre.SetFocus
If txtNombre.Text <> "" Or txtCorreo.Text <> "" Or txtContraseña.Text <> "" Then
Adodc1.Recordset.Update
MsgBox "Se ha guardado correctamente"
Call DesabilitarControles
Else
mensaje = MsgBox("Rellena las casillas", vbCritical, "Todo Jeans")
End If
End Sub
Private Sub cmdM_Click()
    Call HabilitarControles
    cmdM.Enabled = False
End Sub

Private Sub cmdnuevo_Click()
Combo1.Enabled = True
txtNombre.SetFocus
On Error GoTo salida
Adodc1.Recordset.AddNew
MsgBox "Clic a lado del nombre para agregar", vbInformation, "Todo Jeans"
Exit Sub
salida:
MsgBox "Has dando clic dos veces en nuevo registro, tienes que agregar algo", vbCritical, "Todo Jeans"

End Sub

Private Sub Form_Load()
    Call DesabilitarControles
    cmbCuenta.ListIndex = 0
End Sub

Private Sub txtCorreo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        txtContraseña.SetFocus
    End If
End Sub

Private Sub txtNombre_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtCorreo.SetFocus
    End If
End Sub
