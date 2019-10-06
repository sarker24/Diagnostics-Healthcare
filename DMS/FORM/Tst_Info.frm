VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTest_Info 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Information"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12210
   Icon            =   "Tst_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   12210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   9735
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   11895
      Begin VB.TextBox txtS_Code 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   3675
         MaxLength       =   3
         TabIndex        =   0
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtS_Name 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   4305
         TabIndex        =   1
         Top             =   1080
         Width           =   4755
      End
      Begin VB.TextBox txtM_Code 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   255
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtRate 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   9090
         TabIndex        =   2
         Top             =   1080
         Width           =   870
      End
      Begin VB.ComboBox txtType 
         Height          =   315
         ItemData        =   "Tst_Info.frx":000C
         Left            =   9975
         List            =   "Tst_Info.frx":002B
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "txtType"
         Top             =   1050
         Width           =   1605
      End
      Begin VB.ComboBox txtM_Name 
         Height          =   315
         ItemData        =   "Tst_Info.frx":0066
         Left            =   855
         List            =   "Tst_Info.frx":009A
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1050
         Width           =   2805
      End
      Begin VB.CommandButton cmdShow_All 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sho&w All Test Information"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8805
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   2640
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Tst_Info.frx":0150
         Height          =   8295
         Left            =   255
         TabIndex        =   9
         Top             =   1380
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   14631
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BorderStyle     =   0
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
         ColumnCount     =   6
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10245
         TabIndex        =   18
         Top             =   780
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1500
         TabIndex        =   17
         Top             =   780
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3675
         TabIndex        =   16
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5640
         TabIndex        =   14
         Top             =   780
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate (Tk.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9030
         TabIndex        =   13
         Top             =   780
         Width           =   885
      End
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   7680
      Top             =   9960
      Visible         =   0   'False
      Width           =   3930
      _ExtentX        =   6932
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
      Caption         =   "5-SELECT FROM TEST_INFO_SUB"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   5760
      Top             =   10320
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3840
      Top             =   10320
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5760
      Top             =   9960
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "Adodc2"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7890
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10410
      Width           =   1050
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8940
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10410
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10410
      Width           =   1050
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10410
      Width           =   1050
   End
End
Attribute VB_Name = "frmTest_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
        con.connectionstring = strcn.Connection
        con.Open
        Set cmd.ActiveConnection = con
        cmd.CommandText = "exec pro_TEST_INFO_MAIN 'D','" + Trim(txtM_Code.text) + _
        "','" + Trim(txtS_Code.text) + "',''"
        cmd.Execute
        con.Close
        
        '---------------delete from test_info_main-------------------
        'Dim st As String
        Adodc1.connectionstring = strcn.Connection
        Adodc1.RecordSource = "select * from Test_Info_sub where m_code='" & Trim(txtM_Code.text) & "'"
        Adodc1.Refresh
        If Adodc1.Recordset.RecordCount > 0 Then
        Else
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con
            cmd.CommandText = "exec pro_TEST_INFO_MAIN 'D1','" + Trim(txtM_Code) + _
            "','','',''"
            cmd.Execute
            con.Close
        End If
        '-----------------------end----------------------------------
        
        
        
        txtM_Name = ""
        txtM_Name = ""
        txtRate = 0
        txtS_Code = ""
        txtS_Name = ""
        txtType.text = "PATH"
    GetGridData
    End If
End Sub

Private Sub cmdNew_Click()
        txtM_Code = ""
        txtM_Name = ""
        txtM_Name = ""
        txtRate = 0
        txtS_Code = ""
        txtS_Name = ""
        txtType.text = "PATH"
        txtM_Code.SetFocus
        GetGridData
End Sub
Private Sub cmdSave_Click()
    If Len(txtS_Code.text) = 0 Then Exit Sub
    If Len(txtRate.text) = 0 Then Exit Sub
    
    If Trim(txtType.text) = "" Then
        MsgBox "Select Test group"
        txtType.SetFocus
        Exit Sub
    End If
'-----------------------------------------------------------------------------
    Dim st As String
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "select * from Test_Info_main where m_code='" & Trim(txtM_Code.text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
    
        UpdTest_Info_main
    
    '---------SELECT FROM TEST_INFO_SUB----------------
        Adodc5.connectionstring = strcn.Connection
        Adodc5.RecordSource = "select * from Test_Info_sub where m_code='" & Trim(txtM_Code.text) & "' and s_code='" & Trim(txtS_Code.text) & "'"
        Adodc5.Refresh
        If Adodc5.Recordset.RecordCount > 0 Then
            UpdTest_Info_Sub
            UpdTest_Info_Rate
           'InsTest_Info_Rate
    '       cmdUpdate_Click
        Else
            InsTest_Info_Sub
            InsTest_Info_Rate
        End If
    '---------------------------------------------------------
    Else
       InsTest_Info_Main
       InsTest_Info_Sub
       InsTest_Info_Rate
'       cmdUpdate_Click
    End If
       
    txtRate = 0
    txtS_Code = ""
    txtS_Name = ""
    txtType.text = "PATH"
    txtS_Code.SetFocus
     
    GetGridData
    

End Sub
Private Sub InsTest_Info_Main()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_TEST_INFO_MAIN 'I','" + txtM_Code + _
    "','" + txtM_Name + "','" + u_id + "'"
    cmd.Execute
    con.Close
End Sub
Private Sub UpdTest_Info_main()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_TEST_INFO_MAIN 'U','" + txtM_Code + _
    "','" + txtM_Name + "','" + u_id + "'"
    cmd.Execute
    con.Close
End Sub
Private Sub InsTest_Info_Sub()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_TEST_INFO_SUB 'I','" + txtS_Code + _
    "','" + txtS_Name + _
    "','" + txtType + _
    "','" + txtM_Code + _
    "','" + u_id + "'"
    cmd.Execute
    con.Close
    
End Sub
Private Sub UpdTest_Info_Sub()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_TEST_INFO_SUB 'U','" + txtS_Code + _
    "','" + txtS_Name + _
    "','" + txtType + _
    "','" + txtM_Code + _
    "','" + u_id + "'"
    cmd.Execute
    con.Close
    
End Sub

Private Sub InsTest_Info_Rate()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_TEST_INFO_RATE 'I','" + txtM_Code + _
    "','" + txtS_Code + "'," + txtRate + ",'" + u_id + "'"
    cmd.Execute
    con.Close
End Sub
Private Sub UpdTest_Info_Rate()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_TEST_INFO_RATE 'U','" + txtM_Code + _
    "','" + txtS_Code + "'," + txtRate + ",'" + u_id + "'"
    cmd.Execute
    con.Close
End Sub
'Private Sub cmdUpdate_Click()
'
'    Adodc2.connectionstring = strcn.Connection
'    Adodc2.RecordSource = "select m_code,s_code from test_info_rate where m_code='" + txtM_Code + "' and s_code='" + txtS_Code + "'"
'    Adodc2.Refresh
'
'   If Adodc2.Recordset.RecordCount > 0 Then
'        con.connectionstring = strcn.Connection
'        con.Open
'        cmd.CommandText = "delete from test_info_rate where m_code='" + txtM_Code + "' and s_code='" + txtS_Code + "'"
'        cmd.ActiveConnection = con
'        cmd.Execute
'
'        cmd.CommandText = "exec pro_TEST_INFO_RATE 'I','" + txtM_Code + _
'        "','" + txtS_Code + "'," + txtRate + ",'',''"
'        cmd.Execute
'
'   Else
'        con.connectionstring = strcn.Connection
'        con.Open
'        Set cmd.ActiveConnection = con
'        cmd.CommandText = "exec pro_TEST_INFO_RATE 'I','" + txtM_Code + _
'        "','" + txtS_Code + "'," + txtRate + ",'',''"
'        cmd.Execute
'    End If
''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    Adodc4.connectionstring = strcn.Connection
'    Adodc4.RecordSource = "select m_code,s_code from test_info_sub where m_code='" + txtM_Code + "' and s_code='" + txtS_Code + "'"
'    Adodc4.Refresh
'    If Adodc4.Recordset.RecordCount > 0 Then
'        cmd.CommandText = "delete from test_info_sub where m_code='" + txtM_Code + "' and s_code='" + txtS_Code + "'"
'        cmd.ActiveConnection = con
'        cmd.Execute
'
'        cmd.CommandText = "exec pro_TEST_INFO_SUB 'I','" + txtS_Code + _
'        "','" + txtS_Name + "','" + txtType + "','" + txtM_Code + "','',''"
'        cmd.Execute
'
'    Else
'
'        cmd.CommandText = "exec pro_TEST_INFO_SUB 'I','" + txtS_Code + _
'        "','" + txtS_Name + "','" + txtType + "','" + txtM_Code + "','',''"





Private Sub cmdShow_All_Click()
    GetGridDataAll
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
    txtM_Code = DataGrid1.Columns(0)
    txtM_Name = DataGrid1.Columns(1)
    txtS_Code = DataGrid1.Columns(2)
    txtS_Name = DataGrid1.Columns(3)
    txtRate = DataGrid1.Columns(4)
    txtType = DataGrid1.Columns(5)
    
End Sub

'        cmd.Execute
'
'    End If
'''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    con.Close
'End Sub
Private Sub DataGrid1_DblClick()

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
       
    DataGrid1.Columns(0).Width = 374.7402
    DataGrid1.Columns(1).Width = 3179.906
    DataGrid1.Columns(2).Width = 380.0945
    DataGrid1.Columns(3).Width = 4780
    DataGrid1.Columns(4).Width = 800
    DataGrid1.Columns(5).Width = 1580
End Sub

Private Sub txtM_Code_LostFocus()
    Dim st As String
    Adodc1.connectionstring = strcn.Connection
    st = "select * from Test_Info_main where m_code='" & Trim(txtM_Code.text) & "'"
    Adodc1.RecordSource = st
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
    txtM_Name = Adodc1.Recordset!m_name
       
    Else
    
       txtM_Name = ""
       txtS_Code = ""
       txtS_Name = ""
       txtRate = 0
       txtType = "PATH"
       
    End If
    GetGridData

    
End Sub
Private Sub GetGridData()
    Adodc3.connectionstring = strcn.Connection
    Adodc3.RecordSource = "exec pro_Test_Info_FLUSH '1','" + Trim(txtM_Code.text) + "'"
    Adodc3.Refresh
    
    DataGrid1.Columns(0).Width = 374.7402
    DataGrid1.Columns(1).Width = 3179.906
    DataGrid1.Columns(2).Width = 380.0945
    DataGrid1.Columns(3).Width = 4780
    DataGrid1.Columns(4).Width = 800
    DataGrid1.Columns(5).Width = 1580
    
End Sub
Private Sub GetGridDataAll()
    Adodc3.connectionstring = strcn.Connection
    Adodc3.RecordSource = "exec pro_Test_Info_FLUSH '2',''"
    Adodc3.Refresh
    
    DataGrid1.Columns(0).Width = 374.7402
    DataGrid1.Columns(1).Width = 3179.906
    DataGrid1.Columns(2).Width = 380.0945
    DataGrid1.Columns(3).Width = 4780
    DataGrid1.Columns(4).Width = 800
    DataGrid1.Columns(5).Width = 1580
    
End Sub


Private Sub txtRate_Change()
    If Not IsNumeric(txtRate.text) Then
        MsgBox "Only Numaric value allow"
        txtRate = 0
        txtRate.SelStart = 0
        txtRate.SelLength = Len(txtRate)
        txtRate.SetFocus
    End If
End Sub

Private Sub txtRate_GotFocus()
    txtRate.SelStart = 0
    txtRate.SelLength = Len(txtRate)
End Sub
