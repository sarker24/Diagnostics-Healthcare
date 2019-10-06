VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rDoc_Pay 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system"
   ClientHeight    =   2625
   ClientLeft      =   1785
   ClientTop       =   1740
   ClientWidth     =   6915
   Icon            =   "rDoc_Pay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   2535
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6855
      Begin VB.CheckBox Chk_Hide_Tot 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total &Hide"
         Height          =   195
         Left            =   5550
         TabIndex        =   6
         Top             =   570
         Width           =   1095
      End
      Begin VB.CommandButton CmdProcess 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Process"
         Height          =   330
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1950
         Width           =   960
      End
      Begin VB.CommandButton cmdPreview 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pre&view"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1950
         Width           =   960
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Close"
         Height          =   330
         Left            =   5730
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1950
         Width           =   930
      End
      Begin VB.TextBox txtDoc_Name 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4215
      End
      Begin VB.TextBox txtRefer_Code 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1440
         Width           =   1125
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   330
         Top             =   240
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
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
         Caption         =   "2-Doctor Name"
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
         Left            =   3090
         Top             =   240
         Visible         =   0   'False
         Width           =   2280
         _ExtentX        =   4022
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
         Left            =   2190
         TabIndex        =   7
         ToolTipText     =   "Delevary Time"
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   50855938
         UpDown          =   -1  'True
         CurrentDate     =   37163
      End
      Begin MSComCtl2.DTPicker stDt 
         Height          =   285
         Left            =   1230
         TabIndex        =   8
         Top             =   1080
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   50855939
         CurrentDate     =   40958
      End
      Begin MSComCtl2.DTPicker edDT_TM 
         Height          =   285
         Left            =   5100
         TabIndex        =   9
         ToolTipText     =   "Delevary Time"
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   50855938
         UpDown          =   -1  'True
         CurrentDate     =   37163.9993055556
      End
      Begin MSComCtl2.DTPicker edDt 
         Height          =   285
         Left            =   4140
         TabIndex        =   10
         Top             =   1080
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   50855939
         CurrentDate     =   40990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   13
         Top             =   1140
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day between"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1110
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor's ID"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   795
      End
   End
End
Attribute VB_Name = "rDoc_Pay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''----Add For Reporting Perpose----------------------------------------------
Private objReportApp                            As CRPEAuto.Application
Private objReport                               As CRPEAuto.Report
Private objReportDatabase                       As CRPEAuto.Database
Private objReportDatabaseTables                 As CRPEAuto.DatabaseTables
Private objReportDatabaseTable                  As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations        As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                             As CRPEAuto.FormulaFieldDefinition


Private objReportSub                            As CRPEAuto.Report 'sub
Private objReportDatabaseSub                    As CRPEAuto.Database 'sub
Private objReportDatabaseTablesSub              As CRPEAuto.DatabaseTables 'sub
Private objReportDatabaseTableSub               As CRPEAuto.DatabaseTable 'sub
Private objReportFormulaFieldDefinationsSub     As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                          As CRPEAuto.FormulaFieldDefinition


Private ObjPrinterSetting                       As CRPEAuto.PrintWindowOptions
Private rscashmaster                            As ADODB.Recordset
Private rscashdetail                            As ADODB.Recordset
Private Tracer                                  As Integer
Private strGroupName                            As String
Dim temp                                        As Double
Dim temp1                                       As Double
''--------------------------------------------------------------------------------


Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CmdPreview_Click()
    If txtRefer_Code = "" Then
        txtRefer_Code.SetFocus
        Exit Sub
    End If
'    CRViewer1_MODE = 15
'    Viewer.Show vbModal

'    Tracer = 0
    Call PrintReport
    
End Sub

Public Sub PrintReport()

'On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Dim StrPat_ID As String
   
    Set rscashmaster = New ADODB.Recordset
    Set rscashdetail = New ADODB.Recordset
    
'-----------Date formate-------------

'-----------Date formate-------------
If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "exec Rpr_Doc_Pay31 '1'", strcn.Connection
    
    
    
    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

If rscashdetail.State <> 0 Then rscashdetail.Close
        rscashdetail.Open "SELECT Pdate,Refcode,Docname,Payto,Uname,Amount FROM Commission_Pay_Temp", strcn.Connection
                     
    

    
        strPath = App.Path + "\reports\RefDoc121.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        

Set objReportFF = objReportFormulaFieldDefinations.Item(16)
            objReportFF.text = "'" + Format(stDt, "dd-MM-yyyy") + "'"

            Set objReportFF = objReportFormulaFieldDefinations.Item(17)

            objReportFF.text = "'" + Format(edDt, "dd-MM-yyyy") + "'"

''            -------------------Add Discunt------------------
'           If Val(txtTotalDiscount.Text) <> 0 And Val(txtDiscount.Text) <> 0 Then
'           Set objReportFF = objReportFormulaFieldDefinations.Item(3)
'            objReportFF.Text = "'" + parseQuotes(txtDiscount.Text) + " '"
'
'            Set objReportFF = objReportFormulaFieldDefinations.Item(4)
'
'            objReportFF.Text = "'" + parseQuotes(txtTotalDiscount.Text) + " '"
'
'
'            Set objReportFF = objReportFormulaFieldDefinations.Item(5)
'
'            objReportFF.Text = "'" + "Special Discount" + " '"
'
'            Set objReportFF = objReportFormulaFieldDefinations.Item(6)
'
'            objReportFF.Text = "'" + "%" + " '"

'End If
'-------------End Add Discunt-------------------
        objReportDatabaseTable.SetPrivateData 3, rscashmaster
    
    
'For subreport purpose
        Set objReportSub = objReport.OpenSubreport("Doc_Payment_Sub.rpt")
        Set objReportDatabaseSub = objReportSub.Database
        Set objReportDatabaseTablesSub = objReportDatabaseSub.Tables
        Set objReportDatabaseTableSub = objReportDatabaseTablesSub.Item(1)
            objReportDatabaseTableSub.SetPrivateData 3, rscashdetail
'End subreport region
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        If Tracer = 0 Then
        objReport.Preview "Payment Report", , , , , 16777216 Or 524288 Or 65536
        Else
'        objReport.PrintOut
        End If
        
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub

'ErrH:
'
'    Select Case Err.Number
'        Case 20545
'            MsgBox "Request cancelled by the user", vbInformation, "Printing Cancel Information"
'        Case Else
'            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Printing Cancel Information"
'    End Select
    
End Sub

Private Sub cmdProcess_Click()
Call Commissionpay
On Error GoTo err_sub

Dim StrStdt As String
Dim StrSttime As String
Dim StDate_TM As String
Dim strSt_date As String

Dim StrEddt As String
Dim StrEdtime As String
Dim EdDate_TM As String
Dim strEd_date As String

If txtRefer_Code = "" Then
   MsgBox "Reference Code Required"
   txtRefer_Code.SetFocus
   Exit Sub
End If


             StrStdt = Trim(Format(rDoc_Pay.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rDoc_Pay.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
                                   
            '++++++for Ending Date and Time++++++++++++++
             
             StrEddt = Trim(Format(rDoc_Pay.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rDoc_Pay.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
         '   Report15.FormulaFields.Item(17).Text = Chr(34) & Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
         '   Report15.FormulaFields.Item(18).Text = Chr(34) & Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
        
        
          '  strRefer_Code = rDoc_Pay.txtRefer_Code
            strSt_date = Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM")
            strEd_date = Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM")


'------------------------------------------------
    con.connectionstring = strcn.Connection
    con.ConnectionTimeout = 120
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec Rpt_doc_pay2 1,'" & txtRefer_Code & "','" & Format(StDate_TM, "yyyy-mm-dd hh:mm AMPM") & "','" & Format(strEd_date, "yyyy-mm-dd hh:mm AMPM") & "'"
     
    
    Set rs = cmd.Execute
    MsgBox rs!MSG
    con.Close

    cmdPreview.Enabled = True
    cmdPreview.SetFocus
    
'   Call Commissionpay
    
Exit Sub
err_sub:
        MsgBox Err.Description, vbCritical
        Resume Next
    
 
    
End Sub

Private Sub Commissionpay()
On Error GoTo err_sub

Dim StrStdt As String
Dim StrSttime As String
Dim StDate_TM As String
Dim strSt_date As String

Dim StrEddt As String
Dim StrEdtime As String
Dim EdDate_TM As String
Dim strEd_date As String

If txtRefer_Code = "" Then
   MsgBox "Reference Code Required"
   txtRefer_Code.SetFocus
   Exit Sub
End If


             StrStdt = Trim(Format(rDoc_Pay.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rDoc_Pay.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
                                   
            '++++++for Ending Date and Time++++++++++++++
             
             StrEddt = Trim(Format(rDoc_Pay.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rDoc_Pay.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
         '   Report15.FormulaFields.Item(17).Text = Chr(34) & Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
         '   Report15.FormulaFields.Item(18).Text = Chr(34) & Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
        
        
          '  strRefer_Code = rDoc_Pay.txtRefer_Code
            strSt_date = Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM")
            strEd_date = Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM")


'------------------------------------------------
    con.connectionstring = strcn.Connection
    con.ConnectionTimeout = 120
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec Commission_Pay1 1,'" & txtRefer_Code & "','" & Format(StDate_TM, "yyyy-mm-dd hh:mm AMPM") & "','" & Format(strEd_date, "yyyy-mm-dd hh:mm AMPM") & "'"
     
    
    Set rs = cmd.Execute
'    MsgBox rs!MSG
    con.Close

    cmdPreview.Enabled = True
    cmdPreview.SetFocus
    
    
Exit Sub
err_sub:
        MsgBox Err.Description, vbCritical
        Resume Next

End Sub

Private Sub edDt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

End Sub

Private Sub edDT_TM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

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
    stDt.value = Now
'    stDT_TM.value = Now
    
    edDt.value = Now
'    edDT_TM.value = Now
    
    Doc_List_MODE = "rDoc_Pay"
End Sub

Private Sub stDt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

End Sub

Private Sub stDT_TM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

End Sub
Private Sub txtRefer_Code_LostFocus()
On Error GoTo err_sub

    If Len(Trim(txtRefer_Code.text)) = 0 Then Exit Sub
               Adodc2.connectionstring = strcn.Connection
               Adodc2.RecordSource = "exec Pro_FLUSH1 1,'" & Trim(txtRefer_Code.text) & "'"
               Adodc2.Refresh
        
               If Adodc2.Recordset.RecordCount > 0 Then
                   txtDoc_Name.text = Adodc2.Recordset!doc_name
               Else
                   frmDoc_List.Show vbModal
               End If
       
Exit Sub
err_sub:
    MsgBox Err.Description, vbCritical
    Resume Next
    
End Sub
