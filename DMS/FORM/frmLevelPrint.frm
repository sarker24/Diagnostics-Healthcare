VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLevelPrint 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Level Printing System"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmLevelPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   19
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Patient Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15015
      Begin VB.TextBox txtPName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox txtRefdby 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cmbPatientID 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cmbPAge 
         Height          =   315
         Left            =   4920
         TabIndex        =   7
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   8880
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   67174403
         CurrentDate     =   41788
      End
      Begin VB.Label lblPatientID 
         Caption         =   "Patient ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblBillDate 
         Caption         =   "Bill Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblPName 
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblRefdby 
         Caption         =   "Refd. by"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Patient Age"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblSex 
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Investigation Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   15015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   4320
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView List1 
         Height          =   3735
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   6588
         View            =   3
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
         BackColor       =   16761024
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   9000
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdodcSub 
      Height          =   375
      Left            =   7080
      Top             =   9120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc AdodcMain 
      Height          =   375
      Left            =   4440
      Top             =   9120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
End
Attribute VB_Name = "frmLevelPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs                              As New ADODB.Recordset
Private rscashmaster                    As New ADODB.Recordset
Private rsCashDetail                    As ADODB.Recordset
Private rsCashDetail1                   As ADODB.Recordset
Private rsATable                        As New ADODB.Recordset
Private rsAMaster                       As New ADODB.Recordset
Private rsCustomerMaster                As New ADODB.Recordset
Dim str                                 As String
Dim Tracer                              As Integer
Private rsTemp2                         As ADODB.Recordset
Dim flagSlNo                            As Integer
Dim strMood                             As String


Private rsRptRtn                            As ADODB.Recordset
''---------------------------------------------------------------------------
''----Add For Reporting Perpose----------------------------------------------
Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition


Private objReportSub                        As CRPEAuto.Report 'sub
Private objReportDatabaseSub                As CRPEAuto.Database 'sub
Private objReportDatabaseTablesSub          As CRPEAuto.DatabaseTables 'sub
Private objReportDatabaseTableSub           As CRPEAuto.DatabaseTable 'sub
Private objReportFormulaFieldDefinationsSub    As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                         As CRPEAuto.FormulaFieldDefinition


Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
Private rsDailyRpt                          As ADODB.Recordset
'Private Tracer                              As Integer
Private strGroupName                        As String
Dim temp As Double
Dim temp1 As Double
Dim temp2 As Double
''--------------------------------------------------------------------------------

Private Sub cmbPatientID_GotFocus()
cmbPatientID.BackColor = &HFFFFC0
End Sub

Private Sub cmbPatientID_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
 
    AdodcMain.connectionstring = gstrConnection
     Call Srch_Pat_ID
'    AdodcMain.RecordSource = "select pat_id1, pat_name, age, cons, tmp_dt, sex  from pat_info_main where pat_id1='" + cmbPatientID.text + "'"
AdodcMain.RecordSource = "select pat_id1, pat_name, age, cons, tmp_dt, sex  from pat_info_main where pat_id='" + Text1.text + "'"
    AdodcMain.Refresh
    
    If AdodcMain.Recordset.RecordCount > 0 Then
        cmbPatientID.text = AdodcMain.Recordset!pat_id1
        txtPName.text = AdodcMain.Recordset!pat_name
        txtAge.text = AdodcMain.Recordset!age
        txtRefdby.text = AdodcMain.Recordset!cons
        DTPicker1 = AdodcMain.Recordset!tmp_dt
        cmbPAge = AdodcMain.Recordset!Sex
    
    End If
   Call PatDetails
   Call search
End If
End Sub

Private Sub Srch_Pat_ID()

    Dim My_Rst As New ADODB.Recordset
    Dim IntPat_ID As Double
'    cn.ConnectionString = gstrConnection
'    cn.Open
'    Set cmd.ActiveConnection = cn
    
    My_Rst.Open "Select pat_id, pat_id1  from pat_info_main Where pat_id1='" + cmbPatientID.text + "'", cn, adOpenStatic, adLockReadOnly
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!pat_id
  '      MsgBox IntPat_ID
    End If
    Text1.text = IntPat_ID
'    cn.Close
End Sub

Private Sub cmbPatientID_LostFocus()
On Error Resume Next

    If cmbPatientID = "" Then Exit Sub
    If cmbPatientID <> "" Then
        cmbPatientID.TabStop = False
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub Form_Load()
Call Connect
   ModFunction.StartUpPosition Me
   
       List1.ColumnHeaders.Add , , "Name", 5000
        List1.ColumnHeaders.Add , , "Rate"
        List1.ColumnHeaders.Add , , "Group", 1000
        List1.ColumnHeaders.Add , , "Id", 0
   
   
   Call pat_id1
End Sub

Private Sub pat_id1()


    Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT pat_id1 FROM pat_info_main ORDER BY pat_id1 ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbPatientID.AddItem rsTemp2("pat_id1")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub


Private Sub PatDetails()
    AdodcSub.connectionstring = gstrConnection
'    AdodcSub.RecordSource = "exec pro_Test_Info_FLUSH '2',''"
           AdodcSub.RecordSource = "select (select s_name=isnull(s_name,'') from test_info_sub " & _
                                   "Where test_info_sub.s_code = pat_info_sub1.s_code And test_info_sub.m_code = pat_info_sub1.m_code " & _
                                   "and pat_id='" + Text1.text + "') as s_name,test_rate,type from pat_info_sub1   where pat_id='" + Text1.text + "'"
                             
   Set DataGrid1.DataSource = AdodcSub
    AdodcSub.Refresh

End Sub

Private Sub search()
On Error Resume Next
Dim strSQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim i As Integer
    
        
         rs.Open "select (select s_name=isnull(s_name,'') from test_info_sub " & _
                         "Where test_info_sub.s_code = pat_info_sub1.s_code And test_info_sub.m_code = pat_info_sub1.m_code " & _
                         "and pat_id='" + Text1.text + "') as s_name,test_rate,type from pat_info_sub1   where pat_id='" + Text1.text + "'", cn, adOpenStatic, adLockReadOnly
        
        Me.List1.ListItems.Clear
   
    If rs.RecordCount <> 0 Then

          Do Until rs.EOF
               With List1.ListItems.Add
                        .text = rs("s_name")
                        .SubItems(1) = rs("test_rate")
                        .SubItems(2) = rs("type")
'                        .SubItems(3) = rs("DOB")
              End With
           rs.MoveNext

        Loop

    End If
    rs.Close
        
End Sub

Private Sub cmdPrint_Click()
Dim i As Integer
cn.Execute "Delete from Level"
    While i <> List1.ListItems.Count
      

            i = i + 1
            
             If List1.ListItems(i).Checked = True Then
                With List1.ListItems.Item(i)
                   cn.Execute "insert into Level(Name, Rate, [Group], Id) " & _
                              " Values('" & .text & "'," & .SubItems(1) & ",'" & .SubItems(2) & "','" + Text1.text + "') "
                  
    End With
       
     End If
    Wend
    
    Call FetchData
    Call PrintReport

End Sub

Public Function FetchData()
    Set rsRptRtn = New ADODB.Recordset
    
    
    rsRptRtn.Open "select pm.pat_id1, pm.pat_name, pm.age, pm.cons, pm.tmp_dt, pm.sex, lv.Name, " & _
                  "lv.Rate,lv.[Group]  from pat_info_main pm, [Level] lv where pm.pat_id='" + Text1.text + "'", cn, adOpenStatic, adLockReadOnly
    
End Function

Public Sub PrintReport()
    If rsRptRtn.RecordCount < 0 Then
        MsgBox "No Quantity Returned Yet ", vbInformation
        Exit Sub
    End If
    
    Set objReportApp = CreateObject("Crystal.CRPE.Application")
    Set objReport = objReportApp.OpenReport(App.Path & "\reports\Level Tag.rpt")
    
    Set objReportDatabase = objReport.Database
    Set objReportDatabaseTables = objReportDatabase.Tables
    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = objReport.PrintWindowOptions

    ObjPrinterSetting.HasPrintSetupButton = True
    ObjPrinterSetting.HasRefreshButton = True
    ObjPrinterSetting.HasSearchButton = True
    ObjPrinterSetting.HasZoomControl = True

    objReportDatabaseTable.SetPrivateData 3, rsRptRtn
    objReport.DiscardSavedData
'    objReport.Preview "Level Printing", , , , , 16777216 Or 524288 Or 65536
    objReport.PrintOut (False)
'    objReport.PrintingStatus = True
    
    Set objReport = Nothing
    Set objReportDatabase = Nothing
    Set objReportDatabaseTables = Nothing
    Set objReportDatabaseTable = Nothing
    
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Bank Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Bank Information Report"
    End Select
    
End Sub

Private Sub CmdPreview_Click()
Dim i As Integer
cn.Execute "Delete from Level"
    While i <> List1.ListItems.Count
      

            i = i + 1
            
             If List1.ListItems(i).Checked = True Then
                With List1.ListItems.Item(i)
                   cn.Execute "insert into Level(Name, Rate, [Group], Id) " & _
                              " Values('" & .text & "'," & .SubItems(1) & ",'" & .SubItems(2) & "','" + Text1.text + "') "
                  
    End With
       
     End If
    Wend
    
    Call FetchData
    Call previewReport

End Sub

'Public Function FetchData()
'    Set rsRptRtn = New ADODB.Recordset
'
'
'    rsRptRtn.Open "select pm.pat_id1, pm.pat_name, pm.age, pm.cons, pm.tmp_dt, pm.sex, lv.Name, " & _
'                  "lv.Rate,lv.[Group]  from pat_info_main pm, [Level] lv where pm.pat_id='" + Text1.text + "'", cn, adOpenStatic, adLockReadOnly
'
'End Function

Public Sub previewReport()
    If rsRptRtn.RecordCount < 0 Then
        MsgBox "No Quantity Returned Yet ", vbInformation
        Exit Sub
    End If
    
    Set objReportApp = CreateObject("Crystal.CRPE.Application")
    Set objReport = objReportApp.OpenReport(App.Path & "\reports\Level Tag.rpt")
    
    Set objReportDatabase = objReport.Database
    Set objReportDatabaseTables = objReportDatabase.Tables
    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = objReport.PrintWindowOptions

    ObjPrinterSetting.HasPrintSetupButton = True
    ObjPrinterSetting.HasRefreshButton = True
    ObjPrinterSetting.HasSearchButton = True
    ObjPrinterSetting.HasZoomControl = True

    objReportDatabaseTable.SetPrivateData 3, rsRptRtn
    objReport.DiscardSavedData
    objReport.Preview "Level Printing", , , , , 16777216 Or 524288 Or 65536
    
    Set objReport = Nothing
    Set objReportDatabase = Nothing
    Set objReportDatabaseTables = Nothing
    Set objReportDatabaseTable = Nothing
    
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Level Print Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Level Print Report"
    End Select
    
End Sub




