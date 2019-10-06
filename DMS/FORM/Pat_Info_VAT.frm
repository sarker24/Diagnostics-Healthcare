VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPat_Info_VAT 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   FillColor       =   &H007DABD0&
   Icon            =   "Pat_Info_VAT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   12600
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0B4A9&
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   12375
      Begin VB.CommandButton cmdShow 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Show"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Show from main table( datewise)"
         Top             =   1350
         Width           =   1320
      End
      Begin VB.CommandButton cmdDelete_All 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Delete All"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   9060
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Delete all information from VAT table"
         Top             =   1350
         Width           =   1320
      End
      Begin VB.CommandButton cmdShow_VAT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show from New"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   10380
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Show from main table( datewise)"
         Top             =   1350
         Width           =   1830
      End
      Begin VB.CommandButton CmdProcess 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Process"
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
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Press process button for create VAT patient's ID"
         Top             =   270
         Width           =   1770
      End
      Begin VB.CommandButton CmdDaily_State 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dai&ly Statement"
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
         Left            =   2010
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Press process button for create VAT patient's ID"
         Top             =   270
         Width           =   1770
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   10320
         Top             =   240
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
         Left            =   10320
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSComCtl2.DTPicker stDT_TM 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         ToolTipText     =   "Delevary Time"
         Top             =   1260
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   56360962
         UpDown          =   -1  'True
         CurrentDate     =   37163.2916666667
      End
      Begin MSComCtl2.DTPicker stDt 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1260
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56360961
         CurrentDate     =   37306
      End
      Begin MSComCtl2.DTPicker edDT_TM 
         Height          =   285
         Left            =   3990
         TabIndex        =   15
         ToolTipText     =   "Delevary Time"
         Top             =   1260
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   56360962
         UpDown          =   -1  'True
         CurrentDate     =   37163.9993055556
      End
      Begin MSComCtl2.DTPicker edDt 
         Height          =   285
         Left            =   3030
         TabIndex        =   16
         Top             =   1260
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56360961
         CurrentDate     =   37337
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Date and Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   18
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Date and Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3030
         TabIndex        =   17
         Top             =   960
         Width           =   1890
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Height          =   5415
      Left            =   6480
      TabIndex        =   4
      Top             =   2040
      Width           =   6015
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   4995
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   8811
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   14803455
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
         Caption         =   "New Patient"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4995
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   8811
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16769007
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
         Caption         =   "New Entry for VAT"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   6255
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4995
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   8811
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   14680031
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
         Caption         =   "Patient Information"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00C0C0C0&
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
      Left            =   11130
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   1320
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pre&view"
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
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   1320
   End
End
Attribute VB_Name = "frmPat_Info_VAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp_Table As New ADODB.Recordset
Dim GrdColVal As String

Private Sub GetGridData_VAT()
    '+++++++for Starting Date and time+++
            Dim StrStdt As String
            Dim StrSttime As String
            Dim StDate_TM As String
    
            StrStdt = Trim(Format(frmPat_Info_VAT.stDt, "yyyy-mm-dd"))
            StrSttime = Trim(Format(frmPat_Info_VAT.stDT_TM, "hh:mm"))
            StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
                                   
            '++++++for Ending Date and Time++++++++++++++
            Dim StrEddt As String
            Dim StrEdtime As String
            Dim EdDate_TM As String
             
            StrEddt = Trim(Format(rPat_Type.edDt, "yyyy-mm-dd"))
            StrEdtime = Trim(Format(rPat_Type.edDT_TM, "hh:mm"))
            EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "exec Select_VAT_Pat 2,'" + StDate_TM + "','" + EdDate_TM + "'"
    Adodc2.Refresh
    
    Set DataGrid3.DataSource = Adodc2.Recordset
    
    DataGrid3.Columns(0).Visible = False
    
    DataGrid3.Columns(1).Width = 1000
    DataGrid3.Columns(2).Width = 1500
    DataGrid3.Columns(3).Width = 2100
'   DataGrid3.Columns(4).Width = 800
    
End Sub
Private Sub GetGridData()
    '+++++++for Starting Date and time+++
            Dim StrStdt As String
            Dim StrSttime As String
            Dim StDate_TM As String
    
            StrStdt = Trim(Format(frmPat_Info_VAT.stDt, "yyyy-mm-dd"))
            StrSttime = Trim(Format(frmPat_Info_VAT.stDT_TM, "hh:mm"))
            StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
                                   
            '++++++for Ending Date and Time++++++++++++++
            Dim StrEddt As String
            Dim StrEdtime As String
            Dim EdDate_TM As String
             
            StrEddt = Trim(Format(rPat_Type.edDt, "yyyy-mm-dd"))
            StrEdtime = Trim(Format(rPat_Type.edDT_TM, "hh:mm"))
            EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "exec Select_VAT_Pat 1,'" + StDate_TM + "','" + EdDate_TM + "'"
    Adodc1.Refresh
    
    Set DataGrid1.DataSource = Adodc1.Recordset
    
    DataGrid1.Columns(0).Visible = False
    DataGrid1.Columns(1).Width = 1000
    DataGrid1.Columns(2).Width = 1500
    DataGrid1.Columns(3).Width = 2100
    DataGrid1.Columns(4).Width = 800
    
End Sub
       
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDaily_State_Click()
CRViewer1_MODE = 47
    Viewer.Show vbModal
End Sub

Private Sub cmdDelete_All_Click()
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete all information from VAT table ?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
        Del_Pat_Main_VAT
        GetGridData_VAT
        Exit Sub
    End If
End Sub

Private Sub CmdPreview_Click()
    CRViewer1_MODE = 17
    Viewer.Show vbModal
End Sub

Private Sub cmdProcess_Click()
On Error GoTo err_sub

    Call Process_Pat_ID
    
    MsgBox "Process Completed ....."
    
    Exit Sub
err_sub:
    MsgBox Err.Description, vbInformation
    Resume Next
    
End Sub

Private Sub cmdShow_Click()
    GetGridData
End Sub

Private Sub cmdShow_VAT_Click()
        
    If cmdShow_VAT.Caption = "Show from VAT" Then
        DataGrid3.Visible = True
        GetGridData_VAT
        cmdShow_VAT.Caption = "Hide VAT Info"
        Exit Sub
    End If
    
    If cmdShow_VAT.Caption = "Hide VAT Info" Then
        DataGrid3.Visible = False
        cmdShow_VAT.Caption = "Show from VAT"
        Exit Sub
    End If
    
End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next
'Dim GrdColVal As Integer
    GrdColVal = DataGrid1.Columns(0).value
    'MsgBox GrdColVal
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    RS.Open "exec Ins_Into_VAT 1,'" + GrdColVal + "'", strcn.Connection
'    RS.Close
    con.Close
    '----------------check--------
    
    Dim Check As Integer
    Check = 0
    If Temp_Table.RecordCount > 0 Then
        Temp_Table.MoveFirst
    
            While Temp_Table.EOF = False
                If Trim(CStr(Temp_Table!pat_id)) = Trim(CStr(DataGrid1.Columns(0).value)) Then
                    Check = 1
                End If
            Temp_Table.MoveNext
            Wend
        
    If Check = 1 Then
        MsgBox "The Paitent Already Exists"
        Check = 0

        Exit Sub
    End If
'    Temp_Table.MoveFirst
End If

'--------------end check-----

    '+++to insert into TEMPORARY RECORDSET from DATAGRID1++
    
        Temp_Table.AddNew
        Temp_Table!pat_id = DataGrid1.Columns(0).value
        Temp_Table!pat_id1 = DataGrid1.Columns(1).value
        Temp_Table!pat_name = DataGrid1.Columns(2).value
        Temp_Table!doctor_name = DataGrid1.Columns(3).value
        Temp_Table!amount = DataGrid1.Columns(4).value
        DataGrid2.Refresh
'++++++++++for count TOTAL_RATE from Temp_Table+++++++++

GetGridData_VAT 'for datagrid3
DataGrid2.Columns(0).Visible = False

End Sub

Private Sub DataGrid2_Click()
    On Error Resume Next
    If DataGrid2.Columns(0).value = "" Then Exit Sub
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    Dim GrdColVal1 As String
    GrdColVal1 = DataGrid2
'    MsgBox GrdColVal1
    RS.Open "exec Ins_Into_VAT 2,'" + GrdColVal1 + "'", strcn.Connection
'    RS.Close
    con.Close

    Temp_Table.Delete
End Sub


Private Sub DataGrid3_Click()
    On Error Resume Next
    If DataGrid3.Columns(0).value = "" Then Exit Sub
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    Dim GrdColVal2 As String
    GrdColVal2 = DataGrid3
    
    RS.Open "exec Ins_Into_VAT 2,'" + GrdColVal2 + "'", strcn.Connection
'    RS.Close
    con.Close
    GetGridData_VAT
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
      
End Sub

Private Sub Form_Load()
    Temp_rst
    
    stDt = Date
'    stDT_TM = Now
    edDt = Date
'    edDT_TM = Now
    GetGridData
End Sub
Public Sub Temp_rst()
    '--------------------------------------------
    Set Temp_Table = New ADODB.Recordset
    With Temp_Table
        .Fields.Append "pat_id", adDouble
        .Fields.Append "pat_id1", adVarChar, 50
        .Fields.Append "pat_name", adVarChar, 50
        .Fields.Append "doctor_name", adVarChar, 50
        .Fields.Append "amount", adDouble
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set DataGrid2.DataSource = Temp_Table
    
    DataGrid2.ReBind
    DataGrid2.Refresh
    
    DataGrid2.Columns(0).Visible = False
    'DataGrid2.Columns(0).Width =
    DataGrid2.Columns(1).Width = 800
    DataGrid2.Columns(2).Width = 2500.071
    DataGrid2.Columns(3).Width = 810.1418
    DataGrid2.Columns(4).Width = 1110.047
    '2505.26
    
End Sub


Private Sub Del_Pat_Main_VAT()
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "delete from pat_info_main_VAT", con
    'If My_Rst.EOF = False Then
    '    txtItem_Code.Text = My_Rst!item_code
    '    txtItem_Name.Text = My_Rst!item_name
    'Else
    '    txtItem_Name.Text = ""
    'End If
    
    con.Close
End Sub

Private Sub Process_Pat_ID()
    
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec U_VAT_ID ", con
    'If My_Rst.EOF = False Then
    '    IntFont = My_Rst!font_type
    'End If
    con.Close
End Sub

