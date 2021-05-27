VERSION 5.00
Begin VB.Form frmInicial 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   5175
   ClientTop       =   3255
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6360
      Top             =   4560
   End
   Begin VB.Label lblVerificando 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Verificando Directorios..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "VERSION 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ArchivosConfiguracion(3) As String
Sub CargarNombresArchivo()
ArchivosConfiguracion(0) = App.Path & "/configuracion.ini"
ArchivosConfiguracion(1) = App.Path & "/tiket.ini"
ArchivosConfiguracion(2) = App.Path & "/url.txt"
ArchivosConfiguracion(3) = App.Path & "/Base/ventas.accdb"

End Sub
Sub VerificarArchivo(Num)
    Dim Resultado As String
    
    Resultado = Dir(ArchivosConfiguracion(Num))
    If Len(Resultado) > 0 Then
        lblVerificando.Caption = lblVerificando.Caption & "OK"
    Else
        MsgBox "El archivo" & ArchivosConfiguracion(Num) & ", No existe verifique", vbCritical, "Error Critico"
        Timer1.Enabled = False
        Unload Me
        
        End If
End Sub
Sub ConsultarDirectorio(Tiempo As Integer)
    Select Case Tiempo
    Case 1:
        lblVerificando.Caption = "Verificando archivo de configuracion..."
        Call VerificarArchivo(0)
    Case 2:
        lblVerificando.Caption = "Verificando archivo de ticket..."
        Call VerificarArchivo(1)
    Case 3:
        lblVerificando.Caption = "Verificando archivo de ubicacion de  Base..."
        Call VerificarArchivo(2)
    Case 4:
        lblVerificando.Caption = "Verificando archivo de base de datos..."
        Call VerificarArchivo(3)
    Case Else
        Timer1.Enabled = False
        Unload Me
        frmLogin.Show
    End Select
End Sub
Private Sub Form_Load()
 Call CargarNombresArchivo

End Sub

Private Sub Timer1_Timer()
    Static Tiempo As Integer
    Tiempo = Tiempo + 1
    Call ConsultarDirectorio(Tiempo)
End Sub
