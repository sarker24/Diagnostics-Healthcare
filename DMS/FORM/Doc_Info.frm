VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDoctor_Info 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctors Informations [Unique Diagnostic Centre]"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "Doc_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Doctors Details Informations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   9615
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Doc_Info.frx":030A
         Height          =   3120
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5503
         _Version        =   393216
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
         ColumnCount     =   8
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2684.977
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1920.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1080
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Doctors Informations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   9615
      Begin VB.ComboBox cmbFax 
         Height          =   315
         Left            =   825
         TabIndex        =   24
         Top             =   1800
         Width           =   3105
      End
      Begin VB.TextBox txtRefer_Code 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   825
         TabIndex        =   11
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtPhone 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   825
         TabIndex        =   10
         Top             =   1455
         Width           =   3105
      End
      Begin VB.TextBox txtDoc_Name 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   825
         TabIndex        =   9
         Top             =   720
         Width           =   8160
      End
      Begin VB.TextBox txtEmail 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1425
         TabIndex        =   8
         Top             =   2175
         Width           =   2505
      End
      Begin VB.TextBox txtAddr 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   825
         TabIndex        =   7
         Top             =   1080
         Width           =   8160
      End
      Begin VB.TextBox txtFax 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4800
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   3105
      End
      Begin MSComCtl2.DTPicker Birth_date 
         Height          =   285
         Left            =   5040
         TabIndex        =   5
         Top             =   1440
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Format          =   66191361
         CurrentDate     =   37374
      End
      Begin MSComCtl2.DTPicker Marriage_date 
         Height          =   285
         Left            =   7620
         TabIndex        =   12
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         Format          =   66191361
         CurrentDate     =   37374
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Doctor"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3960
         TabIndex        =   14
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Marriage"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   6360
         TabIndex        =   13
         Top             =   1440
         Width           =   1350
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3885
      Top             =   6945
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "2Grid"
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
      Left            =   1845
      Top             =   6960
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
      Height          =   450
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
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
      Height          =   450
      Left            =   7740
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
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
      Height          =   450
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
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
      Height          =   450
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   990
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   120
      Top             =   6960
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
   Begin VB.Label Label10 
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
      TabIndex        =   23
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmDoctor_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim Mode As String

Private Sub Birth_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
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
        cmd.CommandText = "exec pro_DOCTOR_INFO 'D','" + txtRefer_Code + _
        "','" + "0" + "','" + "0" + _
        "','" + "0" + "','" + "0" + "','" + "0" + "','" + "" + "','" + "" + "','" + u_id + "'"
        cmd.Execute
        con.Close
    GetGridData
    txtRefer_Code = ""
    txtDoc_Name = ""
    txtAddr = ""
    txtEmail = ""
    txtFax = ""
    txtPhone = ""
    End If
End Sub

Private Sub cmdNew_Click()
txtAddr = ""
txtDoc_Name = ""
txtEmail = ""
txtFax = ""
txtPhone = ""
'txtRefer_Code = ""

txtRefer_Code.SetFocus
End Sub

Private Sub cmdSave_Click()
    If Trim(txtRefer_Code) = "" Then
    MsgBox "Doctor's ID Mandatory"
    txtRefer_Code.SetFocus
    Exit Sub
    End If
    
    If Trim(txtDoc_Name) = "" Then
    MsgBox "Doctor's Name Mandatory"
    txtDoc_Name.SetFocus
    Exit Sub
    End If
    
    Dim st As String
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Doc_SELECT 1,'" & Trim(txtRefer_Code.text) & "'"

    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        Mode = "U"
        InsDoc_Info
    Else
        Mode = "I"
        InsDoc_Info
    End If
    
    GetCommision
    
    txtDoc_Name = ""
    txtAddr = ""
    txtPhone = ""
    txtFax = ""
    txtEmail = ""
    txtRefer_Code.SetFocus
    
    GetGridData
    
End Sub

Private Sub InsDoc_Info()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_DOCTOR_INFO '" + Mode + "','" + txtRefer_Code + _
    "','" + ChkForQuote(txtDoc_Name) + _
    "','" + ChkForQuote(txtAddr) + _
    "','" + ChkForQuote(txtPhone) + _
    "','" + ChkForQuote(txtFax) + _
    "','" + ChkForQuote(txtEmail) + _
    "','" + Format(Birth_date, "yyyy-mm-dd") + _
    "','" + Format(Marriage_date, "yyyy-mm-dd") + _
    "','" + u_id + "'"
    cmd.Execute
    con.Close
'    "','" + u_id + "','" + Left(Combo1.Text, 2) + "'"
End Sub

Private Sub DataGrid1_Click()
    txtRefer_Code = DataGrid1.Columns(0)
    txtDoc_Name = DataGrid1.Columns(1)
    txtAddr = DataGrid1.Columns(2)
    txtPhone = DataGrid1.Columns(3)
    txtFax = DataGrid1.Columns(4)
    txtEmail = DataGrid1.Columns(5)
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
Adodc1.connectionstring = strcn.Connection
Adodc1.RecordSource = "SELECT Emp_ID + ' - ' + Emp_Name As Doctor From Emp_Info WHERE (Title = 'BD') ORDER BY Emp_ID"
Adodc1.Refresh
    
'If Adodc1.Recordset.RecordCount > 0 Then
'   Do Until Adodc1.Recordset.EOF
'      Combo1.AddItem Adodc1.Recordset!Doctor
'   Adodc1.Recordset.MoveNext
'   Loop
'End If

GetGridData
End Sub

Private Sub nbrCommission_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
End Sub

Private Sub Marriage_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
End Sub

Private Sub txtDoc_Name_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
End Sub

Private Sub txtRefer_Code_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
End Sub

Private Sub txtRefer_Code_LostFocus()
    If Trim(txtRefer_Code.text) = 0 Then Exit Sub
    
    Dim st As String
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Doc_SELECT 1,'" & Trim(txtRefer_Code.text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        txtDoc_Name = Adodc1.Recordset!doc_name
        txtAddr = Adodc1.Recordset!addr
        txtPhone = Adodc1.Recordset!phone
        txtFax = Adodc1.Recordset!fax
        txtEmail = Adodc1.Recordset!email
        Birth_date.value = Adodc1.Recordset!Birth_date
        Marriage_date.value = Adodc1.Recordset!Birth_date
    
    Else
        txtDoc_Name = ""
        txtAddr = ""
        txtPhone = ""
        txtFax = ""
        txtEmail = ""
'        Combo1.Text = ""
    End If
End Sub

Private Sub GetGridData()
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec Pro_FLUSH 1,''"

    Adodc2.Refresh
    DataGrid1.Columns(0).Width = 880.0945
    DataGrid1.Columns(1).Width = 2684.977
    DataGrid1.Columns(2).Width = 1395.213
    DataGrid1.Columns(3).Width = 1395.213
    DataGrid1.Columns(4).Width = 1514.835
    DataGrid1.Columns(5).Width = 1514.835
    DataGrid1.Columns(6).Width = 1500
    DataGrid1.Columns(7).Width = 1500
    DataGrid1.Columns(8).Width = 1500
End Sub

Private Sub GetCommision()

Adodc3.connectionstring = strcn.Connection
    Adodc3.RecordSource = "select * from commission_per where refer_code='" + txtRefer_Code + "'"
    Adodc3.Refresh
'    If Adodc3.Recordset.RecordCount > -1 Then
If Adodc3.Recordset.RecordCount = 0 Then
        
'con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con

    cmd.CommandText = "insert into Commission_Per(type,refer_code,comm_per) " & _
                            " Select top 1 'DOPL', " & _
                            "'" + txtRefer_Code + "', " & " Commission_Per.comm_per from Commission_Per where type='DOPL' union " & _
                            " Select top 1 'ECG', " & _
                            "'" + txtRefer_Code + "', " & " Commission_Per.comm_per from Commission_Per where type='ECG' union " & _
                            " Select top 1 'ECHO', " & _
                            "'" + txtRefer_Code + "', " & " Commission_Per.comm_per from Commission_Per where type='ECHO' union " & _
                            " Select top 1 'ENDO', " & _
                            "'" + txtRefer_Code + "', " & " Commission_Per.comm_per from Commission_Per where type='ENDO' union " & _
                            " Select top 1 'HISTO', " & _
                            "'" + txtRefer_Code + "', " & " Commission_Per.comm_per from Commission_Per where type='HISTO' union " & _
                            " Select top 1 'PATH', " & _
                            "'" + txtRefer_Code + "', " & " Commission_Per.comm_per from Commission_Per where type='PATH' union " & _
                            " Select top 1 'SPATH', " & _
                            "'" + txtRefer_Code + "', " & " Commission_Per.comm_per from Commission_Per where type='SPATH' union " & _
                            " select top 1 'USG'," & _
                            "'" + txtRefer_Code + "', " & " Commission_Per.Comm_per from Commission_Per where type='USG' union " & _
                            " Select top 1 'X-RAY'," & _
                            "'" + txtRefer_Code + "', " & " Commission_Per.comm_per from Commission_Per where type='X-RAY'"
            
            
    cmd.Execute
    con.Close

    Else
    
    End If

End Sub

