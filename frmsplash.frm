VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsplash 
   Caption         =   "Loading"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7380
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmsplash.frx":1084A
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   240
      Top             =   3960
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblLoading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading: "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1395
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
With ProgressBar1
    If .Value < 100 Then
        .Value = .Value + 10
        lblLoading.Caption = "Loading: " & .Value & "%"
    Else
        frmLogin.Show
        Unload Me
    End If
End With
End Sub
