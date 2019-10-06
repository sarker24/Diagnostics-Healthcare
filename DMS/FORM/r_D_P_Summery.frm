VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RptTWStatement 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Test & Department wise Statement"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   Icon            =   "r_D_P_Summery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H00808000&
      Caption         =   "&Process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H00C000C0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdTRate 
      BackColor       =   &H00008080&
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin MSComCtl2.DTPicker edDT_TM 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      ToolTipText     =   "Delevary Time"
      Top             =   480
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   64815106
      UpDown          =   -1  'True
      CurrentDate     =   37163.9993055556
   End
   Begin MSComCtl2.DTPicker stDT_TM 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      ToolTipText     =   "Delevary Time"
      Top             =   480
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   64815106
      UpDown          =   -1  'True
      CurrentDate     =   37163
   End
   Begin MSComCtl2.DTPicker edDt 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   64815107
      CurrentDate     =   40579
   End
   Begin MSComCtl2.DTPicker StDt 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   64815107
      CurrentDate     =   40579
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date and Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2880
      TabIndex        =   8
      Top             =   240
      Width           =   2325
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date and Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2400
   End
End
Attribute VB_Name = "RptTWStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPreview_Click()

End Sub

Private Sub cmdProcess_Click()

On Error GoTo err_sub

Dim StrStdt As String
Dim StrSttime As String
Dim StDate_TM As String
Dim strSt_date As String
'
Dim StrEddt As String
Dim StrEdtime As String
Dim EdDate_TM As String
Dim strEd_date As String

'If txtRefer_Code = "" Then
'   MsgBox "Reference Code Required"
'   txtRefer_Code.SetFocus
'   Exit Sub
'End If


             StrStdt = Trim(Format(frmRefDoc.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(frmRefDoc.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
                                   
            '++++++for Ending Date and Time++++++++++++++
             
             StrEddt = Trim(Format(frmRefDoc.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(frmRefDoc.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
        
        
        
            strSt_date = Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM")
            strEd_date = Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM")


'------------------------------------------------
    con.connectionstring = strcn.Connection
    con.ConnectionTimeout = 120
    con.Open
    Set cmd.ActiveConnection = con
    
'    cmd.CommandText = "exec Rpt_doc_pay22 1,'" & Format(StDate_TM, "yyyy-mm-dd") & "','" & Format(strEd_date, "yyyy-mm-dd") & "'"
     
   cmd.CommandText = "exec Rpt_doc_pay22 1,'" & Format(StDate_TM, "yyyy-mm-dd hh:mm AMPM") & "','" & Format(strEd_date, "yyyy-mm-dd hh:mm AMPM") & "'"
 
    Set rs = cmd.Execute
    MsgBox rs!MSG
    con.Close

'    cmdPreview.Enabled = True
'    cmdPreview.SetFocus
    
    
Exit Sub
err_sub:
        MsgBox Err.Description, vbCritical
        Resume Next
End Sub

Private Sub cmdTRate_Click()
Tracer = 0
    Call TestReport
    End Sub
    
Public Sub TestReport()
On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Dim StrPat_ID As String
   
    Set rscashmaster = New ADODB.Recordset

If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "exec rptTest_State_Rate 1,'" & Format(stDt, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(edDt, "yyyy-mm-dd hh:mm AM/PM") & "'", strcn.Connection
    
    
    
    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Rate.RPT"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
    
'-------------End Add Discunt-------------------
        objReportDatabaseTable.SetPrivateData 3, rscashmaster
    
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

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Printing Cancel Information"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Printing Cancel Information"
    End Select
End Sub

Public Sub PrintReport()
On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Dim StrPat_ID As String
   
    Set rscashmaster = New ADODB.Recordset

If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "exec rptTest_State_Rate 1,'" & Format(stDt, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(edDt, "yyyy-mm-dd hh:mm AM/PM") & "'", strcn.Connection
    
    
    
    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Rate.RPT"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        



             
                   
'        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'            objReportFF.Text = "'" + parseQuotes(txtWords.Text) + " '"
'
'            Set objReportFF = objReportFormulaFieldDefinations.Item(2)
'
'            objReportFF.Text = "'" + parseQuotes(txtUserName.Text) + " '"
'
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

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Printing Cancel Information"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Printing Cancel Information"
    End Select
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
stDt.value = Date
edDt.value = Date
End Sub

