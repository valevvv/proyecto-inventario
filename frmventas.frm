VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmventas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Ventas"
   ClientHeight    =   11730
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17805
   Icon            =   "frmventas.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmventas.frx":1084A
   ScaleHeight     =   782
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1187
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsuario 
      DataField       =   "Id_Usuario"
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
      Left            =   2400
      TabIndex        =   13
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtRopa 
      DataField       =   "Id_Ropa"
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
      Left            =   2400
      TabIndex        =   12
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtTotal 
      DataField       =   "Total"
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
      Left            =   2400
      TabIndex        =   11
      Top             =   7920
      Width           =   2175
   End
   Begin VB.TextBox txtcategoria 
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
      Left            =   2400
      TabIndex        =   10
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox txtCantidad 
      DataField       =   "Cantidad"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtFecha 
      DataField       =   "Fecha"
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
      Height          =   405
      Left            =   2400
      TabIndex        =   8
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdeliminar 
      Height          =   735
      Left            =   13560
      Picture         =   "frmventas.frx":2DA53
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10440
      Width           =   1575
   End
   Begin VB.CommandButton cmdC 
      Height          =   735
      Left            =   9360
      Picture         =   "frmventas.frx":2ED8D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10440
      Width           =   1575
   End
   Begin VB.CommandButton cmdG 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      DisabledPicture =   "frmventas.frx":301C0
      Height          =   735
      Left            =   7080
      MaskColor       =   &H80000005&
      Picture         =   "frmventas.frx":33127
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10440
      Width           =   1575
   End
   Begin VB.CommandButton cmdnuevo 
      Height          =   735
      Left            =   11520
      Picture         =   "frmventas.frx":34517
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10440
      Width           =   1575
   End
   Begin VB.TextBox txtPrecio 
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
      Left            =   2400
      TabIndex        =   1
      Top             =   6720
      Width           =   2175
   End
   Begin VB.TextBox txtStock 
      DataField       =   "Total"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   8880
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3495
      Left            =   5280
      TabIndex        =   2
      Top             =   6480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6165
      _Version        =   393216
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   23
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
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
      Height          =   3975
      Left            =   5280
      TabIndex        =   3
      Top             =   2040
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7011
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Id_Usuario"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Id_Ropa"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   19
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Categoria"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   18
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   7800
      Width           =   1215
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
    'rsinventario.MoveFirst
                rsinventario.Fields("cant_disponible") = Val(rsinventario.Fields("cant_disponible")) + Val(txtCantidad.Text)
                rsinventario.Fields("Valoracion") = rsinventario.Fields("Valoracion") + Val(txtPrecio.Text)
                rsinventario.Update
    rsVentas.Delete
    rsVentas.MoveLast
End Sub

Private Sub cmdG_Click()
    rsinventario.MoveFirst
    If txtFecha.Text <> "" Or txtCantidad.Text <> "" Or txtcategoria.Text <> "" Or txtTotal.Text <> "" Or txtRopa.Text <> "" Or txtUsuario.Text <> "" Then
        rsinventario.Find "Id_ropa = '" & rsropa!Id_Ropa & "'"
        If rsinventario.EOF Then
            MsgBox "No se ha encontrado el articulo"
        Else
            If rsinventario.Fields("Cant_disponible") < Val(txtCantidad.Text) Then
                MsgBox rsinventario.Fields("Cant_disponible") & "Cantidad Insuficiente" & Val(txtCantidad.Text)
            Else
                rsVentas.AddNew
                rsVentas("Fecha") = Date
                rsVentas("cantidad") = txtCantidad.Text
                rsVentas("categoria") = txtcategoria.Text
                rsVentas("total") = txtTotal.Text
                rsVentas("Id_usuario") = txtUsuario.Text
                rsVentas("Id_ropa") = txtRopa.Text 'Se acuerdan que no se ponia el id_ropa en la tabla ventas? era porque se olvidaron de poner esta linea .-.
                rsVentas.Update
                rsinventario.Fields("cant_disponible") = Val(rsinventario.Fields("cant_disponible")) - Val(txtCantidad.Text)
                rsinventario.Fields("Valoracion") = Val(rsinventario.Fields("cant_disponible")) * Val(rsropa.Fields("precio"))
                rsinventario.Update
            End If
        End If
        MsgBox "Se ha guardado correctamente"
    Else
        mensaje = MsgBox("Rellena las casillas", vbCritical, "Todo Jeans")
    End If
    
End Sub



Private Sub DataGrid2_DblClick()
    txtFecha.Text = Date
    txtcategoria.Text = rsropa.Fields("categoria")
    txtRopa.Text = rsropa.Fields("Id_ropa")
    txtPrecio.Text = rsropa.Fields("Precio")
    rsinventario.Find "Id_ropa = '" & rsropa.Fields("Id_ropa") & "'"
    txtStock.Text = rsinventario.Fields("Cant_disponible")
End Sub

Private Sub Form_Load()
    tablaropa
    tablaInventario
    tablaventas
    Set DataGrid1.DataSource = rsVentas
    Set DataGrid2.DataSource = rsropa
    formatoropa
    formatoventas
End Sub

Private Sub txtCantidad_Change()
    txtTotal.Text = Val(txtCantidad.Text) * Val(txtPrecio.Text)
    If txtCantidad.Text > txtStock Then
    MsgBox "mucha cantidad "
    txtCantidad.Text = ""
    End If
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

Private Sub txtRopa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtUsuario.SetFocus
End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCantidad.SetFocus
End If
End Sub

Private Sub formatoropa()
DataGrid2.Columns(0).Width = 70
DataGrid2.Columns(1).Width = 150
DataGrid2.Columns(2).Width = 50
DataGrid2.Columns(3).Width = 200
DataGrid2.Columns(4).Width = 50
DataGrid2.Columns(5).Width = 105
End Sub
Private Sub formatoventas()
DataGrid1.Columns(0).Width = 60
DataGrid1.Columns(1).Width = 150
DataGrid1.Columns(2).Width = 60
DataGrid1.Columns(3).Width = 200
DataGrid1.Columns(4).Width = 70
DataGrid1.Columns(5).Width = 70
DataGrid1.Columns(6).Width = 75
End Sub
