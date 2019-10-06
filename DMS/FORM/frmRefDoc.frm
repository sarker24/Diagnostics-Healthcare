VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRefDoc 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Refferred Doctor's Payment Details"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   Icon            =   "frmRefDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton DSummary 
      BackColor       =   &H000040C0&
      Caption         =   "Department Summary"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker edDT_TM 
      Height          =   285
      Left            =   3840
      TabIndex        =   8
      ToolTipText     =   "Delevary Time"
      Top             =   360
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   50593794
      UpDown          =   -1  'True
      CurrentDate     =   37163.9993055556
   End
   Begin MSComCtl2.DTPicker stDT_TM 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      ToolTipText     =   "Delevary Time"
      Top             =   360
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "HH:MM:SS"
      Format          =   50593794
      UpDown          =   -1  'True
      CurrentDate     =   37163
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Pre&view"
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
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdRefSummary 
      BackColor       =   &H00008080&
      Caption         =   "Ref Summary"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton btnShow 
      BackColor       =   &H00008000&
      Caption         =   "Ref Details"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin MSComCtl2.DTPicker edDt 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   50593795
      CurrentDate     =   40579
   End
   Begin MSComCtl2.DTPicker StDt 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   50593795
      CurrentDate     =   40579
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
      TabIndex        =   10
      Top             =   120
      Width           =   2400
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
      TabIndex        =   9
      Top             =   120
      Width           =   2325
   End
End
Attribute VB_Name = "frmRefDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'Private Tracer                                 As Integer
Private strGroupName                            As String
Dim temp                                        As Double
Dim temp1                                       As Double
''--------------------------------------------------------------------------------

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

Private Sub btnShow_Click()
Tracer = 0
    Call PrintReport
End Sub

Public Sub PrintReport()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Dim StrPat_ID As String
   
    Set rscashmaster = New ADODB.Recordset
    
'-----------Date formate-------------

'-----------Date formate-------------
If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "exec Rpr_Doc_Pay31 '1'", strcn.Connection
    
    
    
    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
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

Public Sub printReport1()
On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Dim StrPat_ID As String
   
    Set rscashmaster = New ADODB.Recordset
    

If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "exec Rpr_Doc_Pay31 '1'", strcn.Connection
    
    
    
    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Summary.RPT"
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


Private Sub CmdPreview_Click()
Tracer = 0
    Call printReport1
End Sub

Private Sub Command2_Click()
Tracer = 0
    Call printReport11
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdRefSummary_Click()
Tracer = 0
    Call PrintReportsummary
End Sub

Public Sub PrintReportsummary()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Dim StrPat_ID As String
   
    Set rscashmaster = New ADODB.Recordset
    
    
'    If frmPatient_Info.txtPat_ID = "" Then
'            StrPat_ID = StPat_ID
'          Else
'            StrPat_ID = frmPatient_Info.txtPat_ID
'    End If

If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "exec Rpr_Doc_Pay31 '1'", strcn.Connection
    
    
    
    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    

        strPath = App.Path + "\reports\Summary.RPT"
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


Private Sub cmdTRate_Click()
Tracer = 0
    Call PrintReport
End Sub



Private Sub DSummary_Click()
Tracer = 0
    Call PrintReport12
End Sub

Public Sub PrintReport12()
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

    
        strPath = App.Path + "\reports\Rate_summary.rpt"
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

Public Sub printReport11()
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

Private Sub Form_Load()
stDt.value = Date
edDt.value = Date
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    Unload Me
    End If
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub


