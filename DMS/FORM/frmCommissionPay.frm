VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCommissionPay 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Doctor Commission Payment"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7305
   Icon            =   "frmCommissionPay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6120
      Picture         =   "frmCommissionPay.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4200
      Picture         =   "frmCommissionPay.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   7095
      Begin VB.TextBox txtDocName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox txtTime 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtSerialNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtAmount 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   9
         Text            =   " "
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtReceivedby 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   8
         Text            =   " "
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtUName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3120
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker PayDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65601539
         CurrentDate     =   41908
      End
      Begin VB.Label lblDoctorCode 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Doctor Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblDoctorName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Doctor Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblSerialNo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Serial No"
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
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Payment Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblPTime 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Payment Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblPAmount 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Payment Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblRBy 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Received by"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblPBy 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Paid by"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   1575
      End
      Begin MSForms.ComboBox cmbDoctorCode 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   720
         Width           =   2295
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4048;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2280
      Picture         =   "frmCommissionPay.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   945
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   360
      Picture         =   "frmCommissionPay.frx":172A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   945
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1320
      Picture         =   "frmCommissionPay.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3240
      Picture         =   "frmCommissionPay.frx":28BE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   945
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5160
      Picture         =   "frmCommissionPay.frx":3188
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000001&
      Caption         =   " Doctor Payment Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmCommissionPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsfactory             As ADODB.Recordset
Private strFileName           As String
Private bRecordExists         As Boolean
Private rm                    As New ADODB.Recordset
Private rs                    As New ADODB.Recordset
Dim str As String
Private rsTemp2                        As ADODB.Recordset

'--------------------------------------------------------------
Private oReportApp                        As CRPEAuto.Application
Private oReport                           As CRPEAuto.Report
Private oReportDatabase                   As CRPEAuto.Database
Private oReportDatabaseTables             As CRPEAuto.DatabaseTables
Private oReportDatabaseTable              As CRPEAuto.DatabaseTable
'Private oReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
'Private oReportFF                         As CRPEAuto.FormulaFieldDefinition
Private ObjPrinterSetting                 As CRPEAuto.PrintWindowOptions

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
     idelete = MsgBox("Do you want to delete this record?", vbYesNo)
     If u_id = "md" Then
    If idelete = vbYes Then
  
            cn.Execute "Delete From Commission_Pay Where SerialNo ='" & parseQuotes(txtSerialNo) & "'"
            Call allClear
    
'    MsgBox "Please Call your System Administrator"
    End If
        
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub CmdPreview_Click()
    Call PrintReport
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdNew.Caption = "&New"
    cmdEdit.Caption = "&Edit"
    cmdPreview.Enabled = True
    CmdDelete.Enabled = True
    cmdOpen.Enabled = True
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    txtSerialNo.Enabled = False
    Call allClear
    Call alldisable
    If Not rsfactory.EOF Then FindRecord
End Sub
'
Private Sub cmdNew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdDelete.Enabled = False
        cmdOpen.Enabled = False
        cmdPreview.Enabled = False
        
        Call allClear
        txtUName.text = u_id
        
If rs.State <> 0 Then rs.Close
'           str = "Select ISNULL(max(SerialNo),0) as SerialNo from Commission_Pay"
            str = "Select ISNULL(max(SerialNo),0) as InvNo from Commission_Pay"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo.text = Val(rs!InvNo) + 1
            
        Call allenable
        cmbDoctorCode.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtSerialNo.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdDelete.Enabled = True
                cmdOpen.Enabled = True
                cmdPreview.Enabled = True
'                Check1.Enabled = True
                Call alldisable
                s = txtSerialNo
                rsfactory.Requery
                rsfactory.MoveFirst
                rsfactory.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
'
'    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        cmbDoctorCode.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdDelete.Enabled = False
        cmdPreview.Enabled = False
        cmdOpen.Enabled = False
'        Check1.Enabled = False
ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdDelete.Enabled = True
'                Check1.Enabled = True
        cmdPreview.Enabled = True
        cmdOpen.Enabled = True
                Call alldisable
                rsfactory.Requery

                Dim s As String
                s = cmbDoctorCode
                rsfactory.Find "refer_code='" & parseQuotes(s) & "'"
'                Call search
'                Call countrysearch
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
    frmCPaySearch.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
   If (KeyCode = 13 And Me.ActiveControl.Name <> "txtAddress") Then SendKeys "{TAB}", True
End Sub

Private Sub Form_Load()

    Call Connect
       ModFunction.StartUpPosition Me
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from Commission_Pay", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If
   
    If Not rsfactory.EOF Then FindRecord
    
    txtSerialNo.Enabled = False
    Call Refer_code
    
End Sub

Private Sub Refer_code()

    Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT refer_code FROM Doctor_Info ORDER BY refer_code ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbDoctorCode.AddItem rsTemp2("refer_code")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
'        If Not rsTemp2.EOF Then FindRecord2

End Sub

'----------------------- Doctor Name Related -------------------------------------------------------

Private Sub cmbDoctorCode_Click()

' If KeyAscii = 13 Then

Set rsfactory = New ADODB.Recordset

    If rsfactory.State <> 0 Then rsfactory.Close
       rsfactory.Open "select refer_code,doc_name,addr from Doctor_Info where refer_code ='" & cmbDoctorCode & "' ", cn, adOpenStatic, adLockReadOnly

   If rsfactory.RecordCount > 0 Then
      rsfactory.MoveFirst
    End If

    If Not rsfactory.EOF Then FindRecord2
End Sub

Private Sub cmbDoctorCode_DropDown()
cmbDoctorCode.Refresh
End Sub

Private Sub FindRecord2()
    txtDocName = rsfactory!doc_name
End Sub

'--------------------------End Doctor Name Informations------------------------------


Private Sub allenable()
    txtDocName.Enabled = True
    cmbDoctorCode.Enabled = True
    PayDate.Enabled = True
    txtAmount.Enabled = True
    txtReceivedby.Enabled = True
    txtUName.Enabled = True
    txtTime.Enabled = True
    
End Sub

Private Sub alldisable()
    txtSerialNo.Enabled = False
    txtDocName.Enabled = False
    cmbDoctorCode.Enabled = False
    PayDate.Enabled = False
    txtAmount.Enabled = False
    txtReceivedby.Enabled = False
    txtUName.Enabled = False
    txtTime.Enabled = False
End Sub

Private Sub allClear()
    cmbDoctorCode.text = ""
    txtDocName.text = ""
    PayDate.value = Date
    txtAmount.text = ""
    txtReceivedby.text = ""
    txtUName.text = u_id
    txtTime.text = Time
   
End Sub

Private Function rcupdate() As Boolean
'    On Error GoTo ErrHandler
    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then
    
    cn.Execute "INSERT INTO Commission_Pay(SerialNo,PDate,refer_code,Doc_Name,Amount,Pay_To,UName,strTime) " & _
                   " VALUES ('" & (txtSerialNo) & "','" & Format(PayDate, "dd-mmm-yyyy") & "','" & parseQuotes(cmbDoctorCode) & "', " & _
                   " '" & parseQuotes(txtDocName) & "', " & _
                   " " & Val(txtAmount.text) & "," & _
                   " '" & parseQuotes(txtReceivedby) & "','" & parseQuotes(txtUName) & "','" & txtTime.text & "')"

          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else
    
    cn.Execute "Update Commission_Pay Set PDate='" & Format(PayDate, "dd-mmm-yyyy") & "',refer_code='" & parseQuotes(cmbDoctorCode) & "', " & _
               "Doc_Name='" & parseQuotes(txtDocName) & ",Amount=" & Val(txtAmount.text) & "',Pay_To='" & parseQuotes(txtReceivedby) & "', " & _
               "strTime='" & txtTime.text & "',UName='" & txtUName.text & "' Where SerialNo ='" & txtSerialNo & "'"

        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
'    Exit Sub
    Exit Function

End Function
Public Sub FindRecord()
If Not rsfactory.EOF Then
        txtSerialNo = rsfactory("SerialNo")
        cmbDoctorCode = rsfactory("refer_code")
        txtDocName = rsfactory("Doc_Name")
        PayDate = rsfactory("PDate")
        txtAmount = rsfactory("Amount")
        txtReceivedby = rsfactory("Pay_To")
        txtUName = rsfactory("UName")
        txtTime = rsfactory("strTime")
    End If
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If (cmbDoctorCode.text = "") Then
       MsgBox "Enter Refd. Code"
       cmbDoctorCode.SetFocus
       IsValidRecord = False
       Exit Function
    End If
    
    If (txtAmount.text = "") Then
      MsgBox "Enter Commission Amount"
      txtAmount.SetFocus
      IsValidRecord = False
      Exit Function
    End If
    

    End Function
'.............................................................................

Public Sub PrintReport()
'On Error GoTo ErrorHan
Dim strPath         As String
Dim rsFactProf      As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\PaymentReciept.rpt"

    Set oReportApp = CreateObject("Crystal.CRPE.Application")
    Set oReport = oReportApp.OpenReport(strPath)
    Set oReportDatabase = oReport.Database
    Set oReportDatabaseTables = oReportDatabase.Tables
    Set oReportDatabaseTable = oReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = oReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select Commission_Pay.SerialNo,Commission_Pay.PDate,Commission_Pay.refer_code, " & _
             "  " & _
             "Commission_Pay.Doc_Name,Commission_Pay.Amount,Commission_Pay.strTime " & _
             "from Commission_Pay where " & _
             "Commission_Pay.SerialNo='" & Me.txtSerialNo & "'"

    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    oReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True

'      Set oReportFormulaFieldDefinations = oReport.FormulaFields
'      Set oReportFF = oReportFormulaFieldDefinations.Item(1)
'      oReportFF.text = "'Factory Information'"

oReport.DiscardSavedData
oReport.Preview "Refd. Doctor Infromation of '" & txtDocName.text & "'", , , , , 16777216 Or 524288 Or 65536

End Sub

Public Sub PopulateCnf(StrID As String)
    rsfactory.MoveFirst
    rsfactory.Find "SerialNo=" & parseQuotes(StrID)
    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub

