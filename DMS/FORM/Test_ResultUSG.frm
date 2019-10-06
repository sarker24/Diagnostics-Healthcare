VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest_ResultUSG 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   11010
   ClientLeft      =   -90
   ClientTop       =   -2040
   ClientWidth     =   11655
   DrawWidth       =   2
   Icon            =   "Test_ResultUSG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4020
      TabIndex        =   11
      Top             =   10680
      Visible         =   0   'False
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtRef_Range 
      BorderStyle     =   0  'None
      Height          =   5820
      Left            =   1665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1935
      Width           =   8490
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C&lear"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10125
      Width           =   1050
   End
   Begin VB.TextBox txtTest_Result 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   8490
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9915
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10125
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10125
      Width           =   1050
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6765
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10125
      Width           =   1050
   End
   Begin VB.TextBox txtUnit 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1170
      Width           =   8490
   End
   Begin VB.TextBox txtTest_Name 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   6570
   End
   Begin VB.TextBox txtType 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1665
      MaxLength       =   4
      TabIndex        =   0
      Top             =   105
      Width           =   1410
   End
   Begin VB.TextBox txtUnique_ID 
      Height          =   285
      Left            =   3090
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   1785
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Test_ResultUSG.frx":000C
      Height          =   2190
      Left            =   870
      TabIndex        =   10
      Top             =   7830
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   3863
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   15725562
      BorderStyle     =   0
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
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2865.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2550.047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2459.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9090
      Top             =   345
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "2-grid"
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
      Left            =   9090
      Top             =   -15
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   330
      TabIndex        =   13
      Top             =   540
      Width           =   1380
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00404000&
      Height          =   195
      Left            =   330
      TabIndex        =   12
      Top             =   135
      Width           =   1380
   End
End
Attribute VB_Name = "frmTest_ResultUSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub InsTest_Result()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_TEST_RESULT 'I','" + Trim(ChkForQuote(txtTest_Name)) + _
    "','" + Trim(ChkForQuote(txtTest_Result)) + _
    "','" + Trim(ChkForQuote(txtUnit)) + _
    "','" + Trim(ChkForQuote(txtRef_Range)) + _
    "','" + Trim(txtType) + _
    "','" + "" + _
    "','" + "" + _
    "','" + "" + _
    "','" + "" + "'"
    cmd.Execute
    con.Close
    
End Sub
Private Sub UpdTest_Result()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_TEST_RESULT1 1,'" + ChkForQuote(txtTest_Name.text) + _
    "','" + ChkForQuote(txtTest_Result.text) + _
    "','" + Trim(ChkForQuote(txtUnit)) + _
    "','" + Trim(ChkForQuote(txtRef_Range)) + _
    "','" + Trim(txtType) + _
    "','" + "" + _
    "','" + "" + _
    "','" + "" + _
    "','" + "" + _
    "'," + txtUnique_ID + ""
    cmd.Execute
    con.Close
    
End Sub

Private Sub cmdClear_Click()
    txtTest_Name = ""
    txtTest_Result = ""
    txtUnit = ""
    txtRef_Range = ""
    txtType = ""
    txtUnique_ID = ""
    txtType.SetFocus
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtType) = "" Then Exit Sub
    If Trim(txtTest_Name) = "" Then Exit Sub
    
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
        con.connectionstring = strcn.Connection
        con.Open
        Set cmd.ActiveConnection = con
'        cmd.CommandText = "exec pro_TEST_RESULT 'D','" + Trim(txtType.Text) + _
'        "','" + Trim(txtTest_Name.Text) + _
'        "','" + "0" + "','" + "0" + _
'        "','" + "1" + "'"
'        cmd.Execute
        cmd.CommandText = "delete from test_result where unique_id='" + Trim(Me.txtUnique_ID.text) + "'"
        cmd.Execute
        con.Close
        
        txtTest_Name = ""
        txtTest_Result = ""
        txtUnit = ""
        txtRef_Range = ""
        txtType = ""
        
        GetGridData
        
    End If
    txtUnique_ID.text = ""
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
    If Trim(txtType) = "" Then
        MsgBox "Type mandatory"
        txtType.SetFocus
        Exit Sub
    End If

    If Trim(txtTest_Name) = "" Then
        MsgBox "Test Name mandatory"
        txtTest_Name.SetFocus
        Exit Sub
    End If


    Adodc1.connectionstring = strcn.Connection
'    Adodc1.RecordSource = "select * from test_result where test_name='" & Trim(txtTest_Name.Text) & "' and type='" & Trim(txtType.Text) & "'"
    'Adodc1.RecordSource = "select * from test_result where unique_id='" & Trim(txtUnique_ID.Text) & "'"
    Adodc1.RecordSource = "exec test_result_SELECT3 1,'" & Trim(txtUnique_ID.text) & "'"

    Adodc1.Refresh
    Me.ProgressBar1.Visible = True
    If Adodc1.Recordset.RecordCount > 0 Then
       UpdTest_Result
       'MsgBox "Updated"
    Else
       InsTest_Result
       'MsgBox "Inserted"
    End If
    
'    txtTest_Name = ""
'    txtTest_Result = ""
'    txtUnit = ""
'    txtRef_Range = ""
''    txtType = ""
    txtUnique_ID.text = ""
    txtTest_Name.SetFocus
    
    GetGridData
    ProgressBar1.Visible = False
End Sub
Private Sub DataGrid1_DblClick()
    txtTest_Name = DataGrid1.Columns(0)
    txtTest_Result = DataGrid1.Columns(1)
    txtUnit = DataGrid1.Columns(2)
    txtRef_Range = DataGrid1.Columns(3)
    txtType = DataGrid1.Columns(4)
    txtUnique_ID = DataGrid1.Columns(9)
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    GetGridData
    
End Sub

Private Sub txtTest_Name_LostFocus()
'    Adodc1.connectionstring = strcn.Connection
'    Adodc1.RecordSource = "exec  test_result_SELECT 1,'" & Trim(txtTest_Name.Text) & "','" & Trim(txtType.Text) & "'"
'
'    Adodc1.Refresh
'    If Adodc1.Recordset.RecordCount > 0 Then
''        txtTest_Name = Adodc1.Recordset!test_name
'        txtTest_Result = Adodc1.Recordset!Test_result
'        txtUnit = Adodc1.Recordset!unit
'        txtRef_Range = Adodc1.Recordset!ref_range
''        txtType = Adodc1.Recordset!Type
'
'    Else
''        txtTest_Name = ""
'        txtTest_Result = ""
'        txtUnit = ""
'        txtRef_Range = ""
''        txtType = ""
'    End If
End Sub
Private Sub GetGridData()
    
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec  test_result_SELECT14"
    Adodc2.Refresh
    DataGrid1.Columns(0).Width = 1500.095
    DataGrid1.Columns(1).Width = 2865.26
    DataGrid1.Columns(2).Width = 2550.047
    DataGrid1.Columns(3).Width = 2459.906
    DataGrid1.Columns(4).Width = 494.9292
    DataGrid1.Columns(5).Visible = False
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub


