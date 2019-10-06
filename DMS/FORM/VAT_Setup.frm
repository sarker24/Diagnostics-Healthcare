VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVAT_Setup 
   BackColor       =   &H00D5EAF7&
   Caption         =   "Prime Diagnostic Ltd."
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3465
   Icon            =   "VAT_Setup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox nbrVAT_Per 
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1380
      TabIndex        =   0
      Top             =   690
      Width           =   870
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1980
      Top             =   120
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT Setup"
      BeginProperty Font 
         Name            =   "AddisonLibbySH"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT Percent"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   690
      Width           =   915
   End
End
Attribute VB_Name = "frmVAT_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Flush_VAT_Per()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '19',''", con
    If My_Rst.EOF = False Then
        nbrVAT_Per.Text = My_Rst!vat_per
    End If
    con.Close

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Flush_VAT_Per
End Sub
Private Sub nbrVAT_Per_Change()
    If Not IsNumeric(nbrVAT_Per.Text) Then
        MsgBox "Only Numaric value allow"
        nbrVAT_Per = 0
        nbrVAT_Per.SelStart = 0
        nbrVAT_Per.SelLength = Len(nbrVAT_Per)
        nbrVAT_Per.SetFocus
    End If
End Sub
Private Sub nbrVAT_Per_GotFocus()
    nbrVAT_Per.SelStart = 0
    nbrVAT_Per.SelLength = Len(nbrVAT_Per)
End Sub
Private Sub VAT_U()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec vat_setup_u 1," & nbrVAT_Per & ""
    'Debug.Print cmd.CommandText
    'cmd.Execute
    'con.Close
     Set RS = cmd.Execute
     MsgBox RS!Message, vbInformation
     con.Close

End Sub

Private Sub nbrVAT_Per_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.nbrVAT_Per = "0" Then Exit Sub
        VAT_U
    End If
End Sub
