VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Main Menu"
   ClientHeight    =   12495
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   28125
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":1084A
   ScaleHeight     =   833
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1875
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Height          =   2655
      Left            =   6840
      Picture         =   "frmMain.frx":373897
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      Height          =   2655
      Left            =   11880
      Picture         =   "frmMain.frx":37A8EF
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Height          =   2655
      Left            =   11880
      Picture         =   "frmMain.frx":382561
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Height          =   2655
      Left            =   6840
      Picture         =   "frmMain.frx":388B62
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Height          =   2655
      Left            =   6840
      Picture         =   "frmMain.frx":38E13A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Height          =   2655
      Left            =   1680
      Picture         =   "frmMain.frx":392DB2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Height          =   2655
      Left            =   1680
      Picture         =   "frmMain.frx":3997EA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   3255
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   12120
      Width           =   28125
      _ExtentX        =   49609
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23230
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   1
            Object.Width           =   23230
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   2970
      Left            =   4800
      Picture         =   "frmMain.frx":39FA2F
      Top             =   0
      Width           =   7365
   End
   Begin VB.Image Image2 
      Height          =   11535
      Left            =   19560
      Picture         =   "frmMain.frx":3A85EF
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   9210
   End
   Begin VB.Image Image1 
      Height          =   19245
      Left            =   0
      Picture         =   "frmMain.frx":4271B9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   30315
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuNitem 
         Caption         =   " Producto"
      End
      Begin VB.Menu j 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIList 
         Caption         =   "Ventas"
      End
      Begin VB.Menu r 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "Reportes"
      Begin VB.Menu mnuRPurchase 
         Caption         =   "Inventario"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRSales 
         Caption         =   "Reporte de ventas"
      End
   End
   Begin VB.Menu mnuSSetting 
      Caption         =   "Configuración"
      Begin VB.Menu mnuSetting 
         Caption         =   "Configuración de usuario"
      End
      Begin VB.Menu v 
         Caption         =   "-"
      End
      Begin VB.Menu s 
         Caption         =   "Configuración de Proveedores"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Registro inventario"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer

Private Sub cmdReporte_Click()
rptVentas.Show 1
End Sub

Private Sub Command1_Click()
frmControl.Show
End Sub

Private Sub Command2_Click()
frmventas.Show
End Sub

Private Sub Command3_Click()
RptInventario.Show
End Sub

Private Sub Command4_Click()
rptVentas.Show
End Sub

Private Sub Command5_Click()
frmEditar.Show
End Sub

Private Sub Command6_Click()
frmregistros.Show
End Sub

Private Sub Command7_Click()
frmproveedores.Show
End Sub

Private Sub mnuBackup_Click()
    frmregistros.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHistory_Click()
    frmTabla.Show
    
End Sub

Private Sub mnuIList_Click()
    frmventas.Show
End Sub

Private Sub mnuNItem_Click()
    frmControl.Show
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Reply = MsgBox("Seguro que desea salir?", vbInformation + vbYesNo)
    If Reply = vbYes Then
        Cancel = 0:                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             SaveSetting "Duax", "Eduard", "Dueñas", X + 1
        End
    Else
        Cancel = 1
    End If
End Sub



Private Sub mnuRPurchase_Click()
 
   RptInventario.Show
End Sub

Private Sub mnuRSales_Click()
   rptVentas.Show
End Sub

Private Sub mnuSetting_Click()
    frmEditar.Show
End Sub








