VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTest_Result_Micro 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   DrawWidth       =   2
   Icon            =   "Test_Result_Micro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
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
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3750
      Width           =   1320
   End
   Begin VB.TextBox txtM_Search 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   10620
      MaxLength       =   4
      TabIndex        =   27
      Top             =   3480
      Width           =   570
   End
   Begin VB.CommandButton cmdSave_As 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save &As"
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
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4320
      Width           =   1320
   End
   Begin VB.TextBox txtOthers1 
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   3180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   5610
      Width           =   6450
   End
   Begin VB.CommandButton cmdShowHaema 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show &Haema"
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
      Left            =   10410
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1530
      Width           =   1320
   End
   Begin VB.TextBox txtS_Code 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   10830
      MaxLength       =   4
      TabIndex        =   7
      Top             =   6180
      Width           =   465
   End
   Begin VB.TextBox txtM_Code 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   10260
      MaxLength       =   4
      TabIndex        =   6
      Top             =   6180
      Width           =   465
   End
   Begin VB.TextBox txtUnique_ID 
      Height          =   285
      Left            =   9675
      TabIndex        =   19
      Top             =   135
      Visible         =   0   'False
      Width           =   1785
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3510
      TabIndex        =   18
      Top             =   7980
      Visible         =   0   'False
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   10380
      Top             =   525
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
      Left            =   10380
      Top             =   165
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
   Begin VB.TextBox txtType 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   825
      MaxLength       =   4
      TabIndex        =   0
      Top             =   435
      Width           =   570
   End
   Begin VB.TextBox txtTest_Name 
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   3165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   6450
   End
   Begin VB.TextBox txtUnit 
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   3165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2280
      Width           =   6450
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
      Left            =   6555
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8295
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
      Left            =   8655
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8295
      Width           =   1050
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
      Left            =   9705
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8295
      Width           =   1050
   End
   Begin VB.TextBox txtTest_Result 
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   3165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   960
      Width           =   6450
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
      Left            =   7605
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8295
      Width           =   1050
   End
   Begin VB.TextBox txtRef_Range 
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   3165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3915
      Width           =   6450
   End
   Begin VB.CommandButton cmdShowUSG 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sho&w USG"
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
      Left            =   10410
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      Width           =   1320
   End
   Begin RichTextLib.RichTextBox txtOthers 
      Height          =   615
      Left            =   3150
      TabIndex        =   30
      Top             =   4980
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   1085
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Test_Result_Micro.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Test_Result_Micro.frx":036F
      Height          =   1470
      Left            =   750
      TabIndex        =   29
      Top             =   6480
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   2593
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   15725562
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "Arial"
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Others1"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1830
      TabIndex        =   25
      Top             =   5610
      Width           =   555
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "T E S T  R E S U L T"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8535
      Left            =   150
      TabIndex        =   24
      Top             =   -30
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Others"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1830
      TabIndex        =   23
      Top             =   5010
      Width           =   465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   10890
      TabIndex        =   21
      Top             =   5850
      Width           =   285
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   10320
      TabIndex        =   20
      Top             =   5850
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   750
      TabIndex        =   17
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference Range"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1830
      TabIndex        =   16
      Top             =   3885
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1830
      TabIndex        =   15
      Top             =   2355
      Width           =   285
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Test"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1830
      TabIndex        =   14
      Top             =   120
      Width           =   960
   End
   Begin VB.Label lblOverflow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1890
      TabIndex        =   13
      Top             =   960
      Width           =   450
   End
End
Attribute VB_Name = "frmTest_Result_Micro"
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
    "','" + Trim(ChkForQuote(txtOthers.Text)) + _
    "','" + Trim(ChkForQuote(txtOthers1.Text)) + _
    "','" + Trim(txtM_Code) + _
    "','" + Trim(txtS_Code) + "'"
    cmd.Execute
    con.Close
    
End Sub
Private Sub UpdTest_Result()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_TEST_RESULT1 1,'" + ChkForQuote(txtTest_Name.Text) + _
    "','" + ChkForQuote(txtTest_Result.Text) + _
    "','" + Trim(ChkForQuote(txtUnit)) + _
    "','" + Trim(ChkForQuote(txtRef_Range)) + _
    "','" + Trim(txtType) + _
    "','" + Trim(ChkForQuote(txtOthers.Text)) + _
    "','" + Trim(ChkForQuote(txtOthers1.Text)) + _
    "','" + Trim(txtM_Code) + _
    "','" + Trim(txtS_Code) + _
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
    txtOthers.Text = ""
    txtOthers1 = ""
    txtM_Code = ""
    txtS_Code = ""
    txtUnique_ID = ""
    txtM_Search = ""
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
        cmd.CommandText = "delete from test_result where unique_id='" + Trim(Me.txtUnique_ID.Text) + "'"
        cmd.Execute
        con.Close
        
        txtTest_Name = ""
        txtTest_Result = ""
        txtUnit = ""
        txtRef_Range = ""
        txtType = ""
        
        GetGridData
        
    End If
    txtUnique_ID.Text = ""
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_As_Click()

    txtUnique_ID.Text = ""

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
    Adodc1.RecordSource = "exec test_result_SELECT3 1,'" & Trim(txtUnique_ID.Text) & "'"

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
    txtUnique_ID.Text = ""
    txtTest_Name.SetFocus
    
    'GetGridData
    
    'GetData_Spec
    Me.ProgressBar1.Visible = False

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
    Adodc1.RecordSource = "exec test_result_SELECT3 1,'" & Trim(txtUnique_ID.Text) & "'"

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
    txtUnique_ID.Text = ""
    txtTest_Name.SetFocus
    
   ' GetGridData
    GetData_Spec1
    
    Me.ProgressBar1.Visible = False
End Sub

Private Sub cmdSearch_Click()
    If txtM_Search = "" Then
        MsgBox " Main Code Mandatory"
        txtM_Search.SetFocus
        Exit Sub
    End If
    GetData_Spec
End Sub

Private Sub cmdShowHaema_Click()
    frmTest_ResultHaema.Show vbModal
End Sub

Private Sub cmdShowUSG_Click()
    frmTest_ResultUSG.Show vbModal
End Sub

Private Sub DataGrid1_DblClick()
    txtTest_Name = DataGrid1.Columns(0)
    txtTest_Result = DataGrid1.Columns(1)
    txtUnit = DataGrid1.Columns(2)
    txtRef_Range = DataGrid1.Columns(3)
    txtType = DataGrid1.Columns(4)
    txtOthers.Text = DataGrid1.Columns(5)
    txtOthers1.Text = DataGrid1.Columns(6)
    txtM_Code = DataGrid1.Columns(7)
    txtS_Code = DataGrid1.Columns(8)
    txtUnique_ID = DataGrid1.Columns(9)
    txtM_Search = DataGrid1.Columns(7)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'GetGridData
    GetData_Spec1
End Sub

Private Sub txtOthers_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF5 Then
        Dim NValue As String
        Dim Power As String
        
        NValue = InputBox("Enter the value of n", "Power of n")

        For i = 0 To (Len(NValue) - 1)
            Select Case Mid(NValue, i + 1, 1)
            Case "0"
            Power = Power & PowerOfNone
            Case "1"
            Power = Power & PowerOfOne
            Case "2"
            Power = Power & PowerOfTwo
            Case "3"
            Power = Power & PowerOfThree
            Case "4"
            Power = Power & PowerOfFour
            Case "5"
            Power = Power & PowerOfFive
            Case "6"
            Power = Power & PowerOfSix
            Case "7"
            Power = Power & PowerOfSeven
            Case "8"
            Power = Power & PowerOfEight
            Case "9"
            Power = Power & PowerOfNine
            End Select
        Next i

        txtOthers.Text = txtOthers.Text & Power

End If


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
    Adodc2.RecordSource = "exec  test_result_SELECT2"
    Adodc2.Refresh
    DataGrid1.Columns(0).Width = 1500.095
    DataGrid1.Columns(1).Width = 2865.26
    DataGrid1.Columns(2).Width = 2550.047
    DataGrid1.Columns(3).Width = 2459.906
    DataGrid1.Columns(4).Width = 494.9292
    DataGrid1.Columns(5).Width = 2000
    DataGrid1.Columns(6).Width = 1000
    DataGrid1.Columns(7).Width = 500
    DataGrid1.Columns(8).Width = 500
    'DataGrid1.Columns(5).Visible = False
    'DataGrid1.Columns(6).Visible = False
    
    DataGrid1.Columns(9).Visible = False
    
End Sub
Private Sub GetData_Spec()
    
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec  test_result_SELECT9 '" & txtM_Search.Text & "'"
    Adodc2.Refresh
    DataGrid1.Columns(0).Width = 1500.095
    DataGrid1.Columns(1).Width = 2865.26
    DataGrid1.Columns(2).Width = 2550.047
    DataGrid1.Columns(3).Width = 2459.906
    DataGrid1.Columns(4).Width = 494.9292
    DataGrid1.Columns(5).Width = 2000
    DataGrid1.Columns(6).Width = 1000
    DataGrid1.Columns(7).Width = 500
    DataGrid1.Columns(8).Width = 500
    'DataGrid1.Columns(5).Visible = False
    'DataGrid1.Columns(6).Visible = False
    
    DataGrid1.Columns(9).Visible = False
    
End Sub
Private Sub txtType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub
Private Sub GetData_Spec1()
    
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec  test_result_SELECT9 '" & "07" & "'"
    Adodc2.Refresh
    DataGrid1.Columns(0).Width = 1500.095
    DataGrid1.Columns(1).Width = 2865.26
    DataGrid1.Columns(2).Width = 2550.047
    DataGrid1.Columns(3).Width = 2459.906
    DataGrid1.Columns(4).Width = 494.9292
    DataGrid1.Columns(5).Width = 2000
    DataGrid1.Columns(6).Width = 1000
    DataGrid1.Columns(7).Width = 500
    DataGrid1.Columns(8).Width = 500
    'DataGrid1.Columns(5).Visible = False
    'DataGrid1.Columns(6).Visible = False
    
    DataGrid1.Columns(9).Visible = False
    
End Sub
