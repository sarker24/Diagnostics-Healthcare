VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCommission_Per 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic Management System"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "Commission_Per.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "All Doctor Commission"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   2160
      Width           =   3135
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
      Height          =   420
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5550
      Width           =   990
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
      Height          =   420
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5550
      Width           =   990
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
      Height          =   420
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5550
      Width           =   990
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
      Height          =   420
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5550
      Width           =   990
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Slelect Doctor's Commission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   8295
      Begin VB.ComboBox txtType 
         Height          =   315
         ItemData        =   "Commission_Per.frx":000C
         Left            =   1320
         List            =   "Commission_Per.frx":002B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1605
      End
      Begin VB.TextBox txtDoc_Name 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   750
         Width           =   5730
      End
      Begin VB.TextBox txtRefer_Code 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   5
         Top             =   750
         Width           =   885
      End
      Begin VB.TextBox nbrComm_Per 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   2310
         TabIndex        =   11
         Top             =   1230
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor's ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   705
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commission"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   1230
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Show All Commission Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   8295
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Commission_Per.frx":0066
         Height          =   2385
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   4207
         _Version        =   393216
         AllowUpdate     =   -1  'True
         ColumnHeaders   =   -1  'True
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
            MarqueeStyle    =   4
            RecordSelectors =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdShow_All 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show &All"
      Height          =   315
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1245
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   120
      Top             =   5520
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
      Left            =   120
      Top             =   5880
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2160
      Top             =   5520
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2160
      Top             =   5880
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   75
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmCommission_Per"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
  If Check1.value = 1 Then
     txtRefer_Code.Visible = False
     txtDoc_Name.Visible = False
  Else
     txtRefer_Code.Visible = True
     txtDoc_Name.Visible = True
  End If
  
End Sub

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
        cmd.CommandText = "exec pro_commission_per 'D','" + txtType + "','" + txtRefer_Code + "'," + "0" + ""
        cmd.Execute
        con.Close

'        txtType = ""
        txtS_Code = ""
        txtRefer_Code = ""
        nbrComm_Per = 0
        txtPhone = ""

    End If
    
    GetGridData
    
End Sub

Private Sub cmdNew_Click()
txtType = "PATH"
'txtM_Name = ""
txtRefer_Code = ""
txtDoc_Name = ""
nbrComm_Per = 0
txtType.SetFocus

End Sub
Private Sub cmdSave_Click()
 If Check1.value = 1 Then
 
   If Trim(txtType) = "" Then
        MsgBox "Main Code Required"
        txtType.SetFocus
        Exit Sub
        End If
        
        
   If Trim(nbrComm_Per) = "" Or Trim(nbrComm_Per) = 0 Then
        MsgBox "Commission Mandatory"
        nbrComm_Per.SetFocus
        Exit Sub
    End If
    
    
     Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "select*from commission_per where type='" + txtType + "'"

    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
'       UpdCommission_Per
       con.connectionstring = strcn.Connection
       con.Open
          Set cmd.ActiveConnection = con

             cmd.CommandText = "Update commission_per " & _
                           "Set (CONVERT(money,comm_per)) = '" + nbrComm_Per + "'" & _
                           "where type='" + txtType + "'"
         cmd.Execute
         con.Close
    Else
'       InsCommission_Per

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con

    cmd.CommandText = "insert into commission_per(type,refer_code,comm_per) " & _
                      " Select '" + txtType + "', " & _
                            "Doctor_Info.refer_code, " & _
                            " " + nbrComm_Per + " from Doctor_Info"
    cmd.Execute
    con.Close
    End If
    
    'GetGridData
    Call GetGridSpec1
    
    nbrComm_Per = 0
    txtType.SetFocus


   Else
   
    If Trim(txtType) = "" Then
        MsgBox "Main Code Required"
        txtType.SetFocus
        Exit Sub
        End If
    
    
    If Trim(txtRefer_Code) = "" Then
        MsgBox "Doctor's ID Required"
        txtRefer_Code.SetFocus
        Exit Sub
    End If
    
    If Trim(nbrComm_Per) = "" Or Trim(nbrComm_Per) = 0 Then
        MsgBox "Commission Mandatory"
        nbrComm_Per.SetFocus
        Exit Sub
    End If
        
    'Dim St As String
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Commission_Per_Select 1,'" & Trim(txtType.text) & "','" & Trim(txtRefer_Code.text) & "'"

    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
       UpdCommission_Per
    Else
       InsCommission_Per
    End If
    
    'GetGridData
    Call GetGridSpec
    
    nbrComm_Per = 0
    txtType.SetFocus
 End If
    
End Sub
Private Sub InsCommission_Per()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_commission_per 'I','" + txtType + _
    "','" + txtRefer_Code + _
    "'," + nbrComm_Per + ""
    cmd.Execute
    con.Close
End Sub
Private Sub UpdCommission_Per()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_commission_per 'U','" + txtType + _
    "','" + txtRefer_Code + _
    "'," + nbrComm_Per + ""
    cmd.Execute
    con.Close
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdShow_All_Click()
    GetGridData
End Sub

Private Sub DataGrid1_Click()
On Error Resume Next
    txtType = DataGrid1.Columns(0)
    txtRefer_Code = DataGrid1.Columns(1)
    txtDoc_Name.text = DataGrid1.Columns(2)
    nbrComm_Per = DataGrid1.Columns(3)
    Exit Sub

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
    txtType.text = "PATH"
    nbrComm_Per = "0"
    GetGridData
End Sub

Private Sub nbrCommission_GotFocus()
    nbrCommission.SelStart = 0
    nbrCommission.SelLength = Len(nbrCommission)
End Sub
Private Sub nbrCommission_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub
Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.ForeColor = &HFF0000
    Label2.ForeColor = &HFF0000
    Label4.ForeColor = &HFF0000
    Label5.ForeColor = &HFF0000
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.ForeColor = &HC000&
    Label2.ForeColor = &HC000&
    Label4.ForeColor = &HC000&
    Label5.ForeColor = &HC000&
    
End Sub

Private Sub nbrComm_Per_Change()
    If Not IsNumeric(nbrComm_Per.text) Then
        MsgBox "Only Numaric value allow"
        nbrComm_Per = 0
        nbrComm_Per.SelStart = 0
        nbrComm_Per.SelLength = Len(nbrComm_Per)
        nbrComm_Per.SetFocus
    End If

End Sub

Private Sub nbrComm_Per_GotFocus()
    nbrComm_Per.SelStart = 0
    nbrComm_Per.SelLength = Len(nbrComm_Per)
End Sub

Private Sub txtType_Click()

  Select Case txtType
        Case "PATH"
            nbrComm_Per = "50"
        Case "SPATH"
            nbrComm_Per = "30"
        Case "HISTO"
            nbrComm_Per = "30"
        Case "X-RAY"
            nbrComm_Per = "25"
        Case "ECG"
            nbrComm_Per = "25"
        Case "USG"
            nbrComm_Per = "25"
        Case "ECHO"
            nbrComm_Per = "30"
        Case "ENDO"
            nbrComm_Per = "20"
        Case "DOPL"
            nbrComm_Per = "30"
    End Select
    

End Sub


Private Sub txtRefer_Code_Change()

    Adodc3.connectionstring = strcn.Connection
    Adodc3.RecordSource = " exec Pro_commission_flush 3,'" & Trim(txtRefer_Code.text) & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
    txtDoc_Name = Adodc3.Recordset!doc_name
    Else
    txtDoc_Name = "Invalied ID"
    End If
    
End Sub

Private Sub GetGridData()
    Adodc4.connectionstring = strcn.Connection
    Adodc4.RecordSource = "exec Pro_FLUSH 7,''"

    Adodc4.Refresh
    
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(1).Width = 1000
    DataGrid1.Columns(2).Width = 4000
    DataGrid1.Columns(3).Width = 1700
    DataGrid1.Refresh
End Sub

Private Sub GetGridSpec()
    Adodc4.connectionstring = strcn.Connection
    Adodc4.RecordSource = "exec FLUSH_Com 1,'" & txtRefer_Code.text & "'"

    Adodc4.Refresh
    
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(1).Width = 1000
    DataGrid1.Columns(2).Width = 4000
    DataGrid1.Columns(3).Width = 1700
    DataGrid1.Refresh
End Sub

Private Sub GetGridSpec1()
    Adodc4.connectionstring = strcn.Connection
'    Adodc5.RecordSource = "exec FLUSH_Com 1,'" & txtRefer_Code.text & "'"
    
   Adodc4.RecordSource = "select type Type,refer_code Doctor_ID, " & _
                "Doctor_Name=(select doc_name from doctor_info where doctor_info.refer_code=commission_per.refer_code),comm_per Commission " & _
                "from commission_per where Type='" & txtType.text & "' order by Type"

    Adodc4.Refresh
    
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(1).Width = 1000
    DataGrid1.Columns(2).Width = 4000
    DataGrid1.Columns(3).Width = 1700
    DataGrid1.Refresh
End Sub

Private Sub txtRefer_Code_LostFocus()

    Adodc3.connectionstring = strcn.Connection
    Adodc3.RecordSource = " exec Pro_commission_flush 3,'" & Trim(txtRefer_Code.text) & "'"
    Adodc3.Refresh
    
    If Adodc3.Recordset.RecordCount > 0 Then
    txtDoc_Name = Adodc3.Recordset!doc_name
    Else
    txtRefer_Code = ""
    txtDoc_Name = "Entry Doctor ID"
    End If
    
    GetGridSpec
    
End Sub

