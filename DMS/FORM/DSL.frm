VERSION 5.00
Begin VB.Form frmDSL 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Daffodil Software Ltd."
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   DrawWidth       =   2
   Icon            =   "DSL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Left            =   390
      Top             =   720
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   840
      Picture         =   "DSL.frx":030A
      Stretch         =   -1  'True
      Top             =   570
      Width           =   2985
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   1545
      Left            =   270
      Shape           =   2  'Oval
      Top             =   240
      Width           =   4185
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H0000FF00&
      Height          =   1485
      Left            =   270
      Shape           =   2  'Oval
      Top             =   420
      Width           =   4515
   End
End
Attribute VB_Name = "frmDSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer3.Enabled = True
    Timer3.Interval = 1000
End Sub

Private Sub Timer3_Timer()
    Unload Me
    frmLogIn.Show vbModal
       
End Sub
