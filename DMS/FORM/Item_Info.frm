VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmItem_Info 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   DrawWidth       =   2
   Icon            =   "Item_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow_All 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show &All"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   960
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   390
      TabIndex        =   9
      Top             =   1260
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4110.236
         EndProperty
      EndProperty
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
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3255
      Width           =   930
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3255
      Width           =   930
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&New"
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
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3255
      Width           =   930
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
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3255
      Width           =   930
   End
   Begin VB.TextBox txtItem_Name 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   1050
      Width           =   4080
   End
   Begin VB.TextBox txtItem_Code 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   1050
      Width           =   1245
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2910
      Top             =   60
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   420
      TabIndex        =   8
      Top             =   120
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   420
      TabIndex        =   7
      Top             =   780
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1740
      TabIndex        =   6
      Top             =   780
      Width           =   540
   End
End
Attribute VB_Name = "frmItem_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
If txtItem_Code.Text = "" Then Exit Sub
    Mode = "D"
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
        Item_Info_IUD
        txtItem_Code = ""
        txtItem_Name = ""
    End If
    
End Sub

Private Sub cmdNew_Click()
    txtItem_Code = ""
    txtItem_Name = ""
    txtItem_Code.SetFocus
    
End Sub
Private Sub cmdSave_Click()
    If txtItem_Code = "" Then Exit Sub
    If txtItem_Name = "" Then Exit Sub
'    If Trim(txtEmp_ID.Text) = "" Then Exit Sub
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec pro_name_SELECT '11','" & txtItem_Code & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Mode = "U"
        Item_Info_IUD
        MsgBox "Updated successfully"
    Else
        Mode = "I"
        Item_Info_IUD
        MsgBox "inserted successfully"
    End If
    
    txtItem_Code = ""
    txtItem_Name = ""
    txtItem_Code.SetFocus
    GetGridData
End Sub
Private Sub Item_Info_IUD()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec Item_Info_IUD '" + Mode + "','" + txtItem_Code.Text + _
    "','" + txtItem_Name.Text + "'"
'Debug.Print cmd.CommandText
    cmd.Execute
    con.Close
End Sub
Private Sub cmdShow_All_Click()
    GetGridData
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
    txtItem_Code = DataGrid1.Columns(0).value
    txtItem_Name = DataGrid1.Columns(1).value
    txtItem_Name.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Private Sub txtItem_Code_LostFocus()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '11','" + Me.txtItem_Code + "'", con
    If My_Rst.EOF = False Then
        txtItem_Code.Text = My_Rst!item_code
        txtItem_Name.Text = My_Rst!item_name
    Else
        txtItem_Name.Text = ""
    End If
    
    con.Close
    
End Sub
Private Sub GetGridData()
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec leave_balance 2,'" + Me.txtItem_Code + "'"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Width = 1280
    DataGrid1.Columns(1).Width = 4080
    
    
End Sub
