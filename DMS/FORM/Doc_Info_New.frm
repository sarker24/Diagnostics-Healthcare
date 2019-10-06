VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDoctor_Info_New 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Doctor Entry"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "Doc_Info_New.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
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
      Height          =   3015
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   8175
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2700
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   4763
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
         ColumnCount     =   7
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
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column06 
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
      Height          =   2415
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   8175
      Begin VB.TextBox txtPhone 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         TabIndex        =   12
         Top             =   1455
         Width           =   6555
      End
      Begin VB.TextBox txtDoc_Name 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         TabIndex        =   0
         Top             =   750
         Width           =   6540
      End
      Begin VB.TextBox txtEmail 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4545
         TabIndex        =   11
         Top             =   1815
         Width           =   3105
      End
      Begin VB.TextBox txtAddr 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         TabIndex        =   1
         Top             =   1095
         Width           =   6540
      End
      Begin VB.TextBox txtFax 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         TabIndex        =   10
         Top             =   1815
         Width           =   2670
      End
      Begin VB.TextBox txtPat_ID 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1110
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   390
         Width           =   1200
      End
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
         Left            =   6510
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker Doc_Date 
         Height          =   285
         Left            =   5190
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   64028673
         CurrentDate     =   37455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3870
         TabIndex        =   16
         Top             =   1815
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1095
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   1815
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   390
         Width           =   840
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1920
      Top             =   6360
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
      Left            =   120
      Top             =   6360
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3228
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
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6375
      Width           =   1050
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
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6375
      Width           =   1050
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6375
      Width           =   1050
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
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6375
      Width           =   1050
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "New Doctor Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmDoctor_Info_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
MsgBox "Delete not Allow"
'    Dim Strmsg As String
'    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
'    If Strmsg = vbYes Then
'        con.connectionstring = strcn.Connection
'        con.Open
'        Set cmd.ActiveConnection = con
'        cmd.CommandText = "exec delete_all 1,'" + txtPat_ID + "'"
'        cmd.Execute
'        con.Close
'
'    GetGridData
    
'    End If
    
End Sub
Private Sub cmdNew_Click()
txtPat_ID = ""
txtAddr = ""
txtDoc_Name = ""
txtEmail = ""
txtFax = ""
txtPhone = ""
txtDoc_Name.SetFocus

End Sub
Private Sub cmdSave_Click()

    If Trim(txtDoc_Name) = "" Then
    MsgBox "Doctor's Name Mandatory"
    txtDoc_Name.SetFocus
    Exit Sub
    End If
    
    If NdocMode = "1" Then
        InsDoc_info_new
    End If
    
    If NdocMode = "0" Then
        UpdDoc_Info_New
    End If
    
'    Dim St As String
'    Adodc1.connectionstring = strcn.Connection
'    Adodc1.RecordSource = "select * from Doctor_Info_New where pat_id='" & Trim(txtPat_ID) & "'"
'    Adodc1.Refresh
'
'    If Adodc1.Recordset.RecordCount > 0 Then
'
'       MsgBox "1"
'
'       UpdDoc_Info_New
'
'    Else
'
'       InsDoc_info_new
'
'    End If

    txtDoc_Name.SetFocus

    GetGridData
    frmPatient_Info.txtDoc_Name = frmDoctor_Info_New.txtDoc_Name
    frmPatient_Info.txtDoc_Addr = frmDoctor_Info_New.txtAddr
    
    Unload Me

    frmPatient_Info.txtRefer_Code = ""
    frmPatient_Info.txtM_Code.SetFocus
   
End Sub
Private Sub InsDoc_info_new()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_DOCTOR_INFO_NEW3 " & txtPat_ID.Text & _
    ",'" & txtDoc_Name & _
    "','" & txtAddr & _
    "','" & txtPhone & _
    "','" & txtFax & _
    "','" & txtEmail & _
    "','" & u_id & _
    "','" & Format(Doc_Date.value, "yyyy-mm-dd") & "'"
    cmd.Execute
    con.Close
    
End Sub
Private Sub UpdDoc_Info_New()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_DOCTOR_INFO_NEW 'U'," + txtPat_ID + _
    ",'" + txtDoc_Name + _
    "','" + txtAddr + _
    "','" + txtPhone + _
    "','" + txtFax + _
    "','" + txtEmail + _
    "','" + u_id + _
    "','" + Format(Doc_Date.value, "yyyy-mm-dd") + "'"

    cmd.Execute
    con.Close
End Sub
Private Sub cmdShow_All_Click()
    GetGridData
End Sub

Private Sub DataGrid1_Click()
    'txtPat_ID = DataGrid1.Columns(0)
    txtDoc_Name = DataGrid1.Columns(1)
    txtAddr = DataGrid1.Columns(2)
    txtPhone = DataGrid1.Columns(3)
    txtFax = DataGrid1.Columns(4)
    txtEmail = DataGrid1.Columns(5)
'   Doc_Date.value = DataGrid1.Columns(6).value
   
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
    Doc_Date.value = Date
    'MsgBox "d"
    If frmPatient_Info.txtPat_ID.Text <> "" Then
            frmDoctor_Info_New.txtPat_ID = frmPatient_Info.txtPat_ID
    End If
    
    If frmPatient_Info.txtPat_ID.Text = "" Then
            txtDoc_Name = ""
            txtAddr = ""
            txtPhone = ""
            txtFax = ""
            txtEmail = ""
        
            frmDoctor_Info_New.txtPat_ID = frmPatient_Info.txtRefer_Code
    End If
    
   
 

   
       If NdocMode = "0" Then
            Flush_Doc_Name
       End If
       
       If NdocMode = "1" Then
            Flush_Doc_Name1
       End If
       

End Sub

'Private Sub nbrCommission_Change()
'    If Not IsNumeric(nbrCommission) Then
'        MsgBox "Only Numaric value allow"
'        nbrCommission = 0
'        nbrCommission.SelStart = 0
'        nbrCommission.SelLength = Len(nbrCommission)
'        nbrCommission.SetFocus
'    End If
'End Sub
'Private Sub nbrCommission_GotFocus()
'    nbrCommission.SelStart = 0
'    nbrCommission.SelLength = Len(nbrCommission)
'End Sub
Private Sub nbrCommission_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
End Sub





Private Sub lblName_Click()
lblName.Caption = txtDoc_Name
End Sub

Private Sub txtDoc_Name_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
End Sub

Private Sub txtDoc_Name_LostFocus()
'       If Trim(txtDoc_Name.Text) = 0 Then Exit Sub
'
'    Dim st As String
'    Adodc1.connectionstring = strcn.Connection
'    st = "select * from Doctor_Info_new where doc_name='" & Trim(txtDoc_Name.Text) & "'"
'    Adodc1.RecordSource = st
'    Adodc1.Refresh
'    If Adodc1.Recordset.RecordCount > 0 Then
'        txtDoc_Name = Adodc1.Recordset!doc_name
''        txtDegree = Adodc1.Recordset!degree
'        txtAddr = Adodc1.Recordset!addr
'        txtPhone = Adodc1.Recordset!phone
'        txtFax = Adodc1.Recordset!fax
'        txtEmail = Adodc1.Recordset!email
'    Else
''        txtDoc_Name = ""
''        txtDegree = ""
'        txtAddr = ""
'        txtPhone = ""
'        txtFax = ""
'        txtEmail = ""
''        nbrCommission = 0
'    End If
End Sub
Private Sub GetGridData()
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec Pro_FLUSH 8,''"

    Adodc2.Refresh
    
    DataGrid1.Columns(0).Width = 1050
    DataGrid1.Columns(1).Width = 2684.977
    DataGrid1.Columns(2).Width = 1395.213
    DataGrid1.Columns(3).Width = 1260.284
    DataGrid1.Columns(4).Width = 1230.236
    DataGrid1.Columns(5).Width = 1289.764
    DataGrid1.Columns(6).Width = 950
    'DataGrid1.Columns(7).Width = 950
    'Dim ww As Date
    'ww = DataGrid1.Columns(6).value
End Sub

Private Sub txtPat_ID_Change()
    If Not IsNumeric(txtPat_ID.Text) Then
        'MsgBox "Only Numaric value allow"
        txtPat_ID = ""
        txtPat_ID.SelStart = 0
        txtPat_ID.SelLength = Len(txtPat_ID)
        'txtPat_ID.SetFocus
    End If
End Sub

Private Sub txtPat_ID_GotFocus()
    txtPat_ID.SelStart = 0
    txtPat_ID.SelLength = Len(txtPat_ID)
End Sub
Private Sub Flush_Doc_Name()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Doc_SELECT 3,'" + txtPat_ID + "'", con
    If My_Rst.EOF = False Then
        txtDoc_Name = My_Rst!doc_name
        txtAddr.Text = My_Rst!addr
        txtPhone.Text = My_Rst!phone
        txtFax.Text = My_Rst!fax
        txtEmail.Text = My_Rst!email
    Else
        txtDoc_Name = ""
        txtAddr.Text = ""
        txtPhone.Text = ""
        txtFax.Text = ""
        txtEmail.Text = ""
        
    End If
    con.Close
End Sub
Private Sub Flush_Doc_Name1()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec New_Doc_Select 1,'" & txtPat_ID & "','" & u_id & "'", con
    If My_Rst.EOF = False Then
        txtDoc_Name = My_Rst!doc_name
        txtAddr.Text = My_Rst!addr
        txtPhone.Text = My_Rst!phone
        txtFax.Text = My_Rst!fax
        txtEmail.Text = My_Rst!email
    Else
        txtDoc_Name = ""
        txtAddr.Text = ""
        txtPhone.Text = ""
        txtFax.Text = ""
        txtEmail.Text = ""
        
    End If
    con.Close
End Sub

Private Sub Flush_Doc_New()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Flush_New_Doc 1,'" & txtPat_ID & "','" & u_id & "'", con
    If My_Rst.EOF = False Then
        txtDoc_Name = My_Rst!doc_name
        txtAddr.Text = My_Rst!addr
        txtPhone.Text = My_Rst!phone
        txtFax.Text = My_Rst!fax
        txtEmail.Text = My_Rst!email
    Else
        txtDoc_Name = ""
        txtAddr.Text = ""
        txtPhone.Text = ""
        txtFax.Text = ""
        txtEmail.Text = ""
        
    End If
    con.Close
End Sub
