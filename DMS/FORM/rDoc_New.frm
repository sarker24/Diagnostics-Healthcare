VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rDoc_New 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "rDoc_New.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Specific"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3630
      TabIndex        =   1
      Top             =   840
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "All"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1710
      TabIndex        =   0
      Top             =   840
      Width           =   675
   End
   Begin VB.ComboBox CombDoc_Name 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Top             =   1860
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0DBE8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4950
      ScaleHeight     =   225
      ScaleWidth      =   1785
      TabIndex        =   11
      Top             =   1920
      Width           =   1785
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   330
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   -60
         Width           =   930
      End
      Begin VB.CommandButton cmdPreview 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pre&view"
         Height          =   330
         Left            =   -30
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   -60
         Width           =   960
      End
   End
   Begin MSComCtl2.DTPicker EdDate 
      Height          =   315
      Left            =   3660
      TabIndex        =   3
      Top             =   1380
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62455809
      CurrentDate     =   37188
   End
   Begin MSComCtl2.DTPicker StDate 
      Height          =   285
      Left            =   1710
      TabIndex        =   2
      Top             =   1380
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   62455809
      CurrentDate     =   37188
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5010
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "2-Doctor Name"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3840
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0B4A9&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   2535
      Left            =   30
      Top             =   30
      Width           =   6885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3180
      TabIndex        =   10
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day between"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   540
      TabIndex        =   9
      Top             =   1380
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Name"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   540
      TabIndex        =   8
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Doctor's Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   540
      TabIndex        =   7
      Top             =   270
      Width           =   3060
   End
End
Attribute VB_Name = "rDoc_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CmdPreview_Click()
    If CombDoc_Name.Enabled = True Then
        If CombDoc_Name = "" Then
            CombDoc_Name.SetFocus
            Exit Sub
        End If
    End If
   
    CRViewer1_MODE = 16
    Viewer.Show vbModal
End Sub

Private Sub CombDoc_Name_Click()
    Show_Doc_Name_New
End Sub

Private Sub CombDoc_Name_GotFocus()
    Me.CombDoc_Name = ""
    CombDoc_Name.Refresh
'    Show_Doc_Name_New
End Sub

Private Sub EdDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub
Private Sub Show_Doc_Name_New()
'    CombDoc_Name = ""
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Select_New_Doc_Name 1,'" + Format(StDate.value, "yyyy-mm-dd") + "','" + Format(EdDate.value, "yyyy-mm-dd") + "'"
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
'       CombDoc_Name = ""
       Do Until Adodc1.Recordset.EOF
'         CombDoc_Name.AddItem Adodc1.Refresh
         
         CombDoc_Name.AddItem Adodc1.Recordset!doc_name
       
       Adodc1.Recordset.MoveNext
       Loop
    End If
End Sub

Private Sub Form_Load()
    StDate = Date
    EdDate = Date
    Show_Doc_Name_New
End Sub

Private Sub Option1_Click()
    If Option1.value = True Then
        CombDoc_Name = ""
        CombDoc_Name.Enabled = False
    End If
End Sub
Private Sub Option2_Click()
    If Option2.value = True Then
       CombDoc_Name.Enabled = True
    End If
End Sub

Private Sub StDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

End Sub
