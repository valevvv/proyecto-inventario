VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H000000FF&
   Caption         =   "Login"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "frmlogin.frx":1084A
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4920
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Picture         =   "frmlogin.frx":144D7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Picture         =   "frmlogin.frx":15618
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   3480
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   582
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
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With Adodc1.Recordset
.Filter = "Nombre = '" & Text1.Text & "' and Contraseña = '" & Text2.Text & "'"
If .RecordCount > 0 Then
    MsgBox "BIENVENIDO: " & .Fields("TipoCuenta"), vbInformation, "Login"
    frmMain.Show
    Unload Me
Else
    MsgBox "Cuenta Invalida", vbCritical, "Login" '
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
End If
End With
End Sub

Private Sub Command2_Click()
End
End Sub

