VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFont 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   DrawWidth       =   2
   Icon            =   "Font.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComFont_Name 
      Height          =   315
      ItemData        =   "Font.frx":000C
      Left            =   3390
      List            =   "Font.frx":0019
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   930
      Width           =   3045
   End
   Begin VB.ComboBox ComScreen_Name 
      Height          =   315
      ItemData        =   "Font.frx":0052
      Left            =   360
      List            =   "Font.frx":008F
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   930
      Width           =   3045
   End
   Begin VB.TextBox nbrFont_Type 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2925
      Width           =   900
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2925
      Width           =   900
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   360
      TabIndex        =   5
      Top             =   1260
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   13099768
      ColumnHeaders   =   0   'False
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
            LCID            =   1033
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
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004.788
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2370
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Name"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Name"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3510
      TabIndex        =   7
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Setup"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   390
      TabIndex        =   6
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If ComScreen_Name = "" Then Exit Sub
    If ComFont_Name = "" Then Exit Sub
    
    Font_IU
    Show_Grid
End Sub
Private Sub ComFont_Name_Click()

    If ComFont_Name.Text = "XavierPlatoSH" Then
        nbrFont_Type = 3
    End If

    If ComFont_Name.Text = "Times New Roman (Western)" Then
        nbrFont_Type = 2
    End If
      
    If ComFont_Name.Text = "YuriKaySH" Then
        nbrFont_Type = 4
    End If
    
End Sub
Private Sub Font_IU()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec font_IU '" & ComScreen_Name & _
    "','" & ComFont_Name & _
    "'," & nbrFont_Type & ""
    'Debug.Print cmd.CommandText
    'cmd.Execute
    'con.Close
     Set RS = cmd.Execute
     MsgBox RS!Message, vbInformation
     con.Close

End Sub
Private Sub Show_Grid()

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec pro_name_SELECT '17',''"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Columns(0).Width = 3040
    DataGrid1.Columns(1).Width = 3020
    DataGrid1.Columns(2).Visible = False
    DataGrid1.Columns(3).Visible = False
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Show_Grid
End Sub
