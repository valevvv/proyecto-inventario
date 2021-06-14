VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmventas 
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3855
      Left            =   7800
      TabIndex        =   18
      Top             =   840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6800
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
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "ELIMINAR"
      Height          =   855
      Left            =   7560
      TabIndex        =   17
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "CERRAR"
      Height          =   855
      Left            =   4800
      TabIndex        =   16
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdG 
      Caption         =   "GUARDAR"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "NUEVO"
      Height          =   855
      Left            =   2400
      TabIndex        =   14
      Top             =   5400
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1440
      Top             =   6480
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
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":0092
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Ventas"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0124
      Height          =   3975
      Left            =   4200
      TabIndex        =   13
      Top             =   840
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7011
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
   Begin VB.TextBox txtUsuario 
      DataField       =   "Id_Usuario"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   1680
      TabIndex        =   12
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txtRopa 
      DataField       =   "Id_Ropa"
      Height          =   615
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtTotal 
      DataField       =   "Total"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   1680
      TabIndex        =   10
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtCategoria 
      DataField       =   "Categoria"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtCantidad 
      DataField       =   "Cantidad"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtFecha 
      DataField       =   "Fecha"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Id_Usuario"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Id_Ropa"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Total"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Categoria"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "VENTAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "frmventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdC_Click()
Unload Me
End Sub

Private Sub cmdeliminar_Click()
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
   Adodc1.Recordset.MoveLast
 End If
End Sub

Private Sub cmdG_Click()
txtFecha.SetFocus
If txtFecha.Text <> "" Or txtCantidad.Text <> "" Or txtCategoria.Text <> "" Or txtTotal.Text <> "" Or txtRopa.Text <> "" Or txtUsuario.Text <> "" Then
Adodc1.Recordset.Update
MsgBox "Se ha guardado correctamente"
Else
mensaje = MsgBox("rellena las casillas", vbCritical, "SISTEMA TIENDA")
End If
End Sub

Private Sub cmdnuevo_Click()
txtFecha.SetFocus
On Error GoTo salida
Adodc1.Recordset.AddNew
MsgBox "clic a lado del nombre para agregar", vbInformation, "SITEMA TIENDA"
Exit Sub
salida:
MsgBox "Has dando clic dos veces en nuevo registro tiene que agregar algo", vbCritical, "SISTEMA TIENDA"

End Sub

Private Sub Form_Load()
Set DataGrid2.DataSource = rsropa
End Sub

Private Sub txtRopa_Change()
Dim buscar As String
buscar = txtRopa.Text & "%"
If rsropa.State = 1 Then rsropa.Close
   rsropa.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rsropa.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    rsropa.Open "select * from ropa where Id_ropa like '%" & buscar & "'", con
Set DataGrid2.DataSource = rsropa
End Sub
