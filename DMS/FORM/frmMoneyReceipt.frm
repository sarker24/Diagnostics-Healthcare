VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMoneyReceipt 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Voucher Entry"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   Icon            =   "frmMoneyReceipt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7920
      Top             =   5040
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SMS;Data Source=NOTEBOOK"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SMS;Data Source=NOTEBOOK"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DCSearch"
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
   Begin VB.Timer Timer1 
      Left            =   10080
      Top             =   5880
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Find First"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Find Previous"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Find Last"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Find Next"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdPost 
      BackColor       =   &H00C0B4A9&
      Caption         =   "P&ost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7560
      MouseIcon       =   "frmMoneyReceipt.frx":000C
      Picture         =   "frmMoneyReceipt.frx":06F6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   6480
      Picture         =   "frmMoneyReceipt.frx":0DE0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5520
      Picture         =   "frmMoneyReceipt.frx":16AA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5520
      Width           =   990
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4440
      Picture         =   "frmMoneyReceipt.frx":1F74
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3360
      Picture         =   "frmMoneyReceipt.frx":283E
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2280
      Picture         =   "frmMoneyReceipt.frx":3108
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1200
      Picture         =   "frmMoneyReceipt.frx":39D2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      Picture         =   "frmMoneyReceipt.frx":429C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Width           =   10455
      Begin VB.TextBox txtDescription 
         Height          =   2025
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Text            =   " "
         Top             =   360
         Width           =   10215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   2295
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10455
      Begin VB.ComboBox cmbAMode 
         Height          =   315
         ItemData        =   "frmMoneyReceipt.frx":4B66
         Left            =   3600
         List            =   "frmMoneyReceipt.frx":4B68
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cmbAHName 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   8880
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtCPost 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "frmMoneyReceipt.frx":4B6A
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtUName 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtDebit 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   375
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtVID 
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
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   375
         Left            =   7320
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker VDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   4210688
         CalendarTitleBackColor=   8421376
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   61603843
         CurrentDate     =   39739
      End
      Begin VB.Label lblPaymentMode 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Account Mode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   30
         Top             =   1320
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8880
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblAHead 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Head of Accounts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblCredit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblDebit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Debit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblVID 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Voucher No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMoneyReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsCashDetail          As ADODB.Recordset
Private rsAMode               As New ADODB.Recordset
Private rsfactory             As ADODB.Recordset
Private strFileName           As String
Private bRecordExists         As Boolean
Private rsTemp2               As ADODB.Recordset
Private rm                    As New ADODB.Recordset
Private rs                    As New ADODB.Recordset
Dim Tracer                    As Integer
Dim strMood                   As String

Dim str As String
'--------------------------------------------------------------
Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private rsDailyRpt                          As ADODB.Recordset
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private ObjPrinterSetting                 As CRPEAuto.PrintWindowOptions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition






Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.EOF = True Then
       cmdFirst.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

        txtVID = Adodc1.Recordset!VID
        VDate.value = Adodc1.Recordset!VDate
        cmbAHName = Adodc1.Recordset!AHead
        txtDebit = Adodc1.Recordset!Debit
        txtCredit = Adodc1.Recordset!Credit
        cmbAMode = Adodc1.Recordset!AMode
        txtDescription = Adodc1.Recordset!Description
        txtTime = Adodc1.Recordset!StrTIME
        txtUName = Adodc1.Recordset!UName
        txtCPost = Adodc1.Recordset!Posted
End If


'End If
End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdLast.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtVID = Adodc1.Recordset!VID
        VDate.value = Adodc1.Recordset!VDate
        cmbAHName = Adodc1.Recordset!AHead
        txtDebit = Adodc1.Recordset!Debit
        txtCredit = Adodc1.Recordset!Credit
        cmbAMode = Adodc1.Recordset!AMode
        txtDescription = Adodc1.Recordset!Description
        txtTime = Adodc1.Recordset!StrTIME
        txtUName = Adodc1.Recordset!UName
        txtCPost = Adodc1.Recordset!Posted

End If
End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdNext.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtVID = Adodc1.Recordset!VID
        VDate.value = Adodc1.Recordset!VDate
        cmbAHName = Adodc1.Recordset!AHead
        txtDebit = Adodc1.Recordset!Debit
        txtCredit = Adodc1.Recordset!Credit
        cmbAMode = Adodc1.Recordset!AMode
        txtDescription = Adodc1.Recordset!Description
        txtTime = Adodc1.Recordset!StrTIME
        txtUName = Adodc1.Recordset!UName
        txtCPost = Adodc1.Recordset!Posted

End If
End Sub

Private Sub cmdPost_Click()
Dim s As String

cmdPost.Caption = "&Posted"

If cmdPost.Caption = "&Posted" Then
     If txtCPost.text = "Not Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                 cmdNew.Caption = "&New"
                 cmdEdit.Enabled = True
                 CmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 cmdFind.Enabled = False
                 cmdPreview.Enabled = True
                 cmdPrint.Enabled = True
                 txtVID.Enabled = False
                 Call alldisable
           End If
        End If
      End If
Else
 End If
cmdPost.Caption = "&Posted"

End Sub

Private Sub CmdPreview_Click()
'    Call printReport
 Tracer = 0
'    Call printReport
If txtCPost.text = "Posted" Then
  If txtCredit.text = "0" Then
   Call Debit

   Else
   Call Credit
   End If
   End If

End Sub


Private Sub cmbAHName_GotFocus()
cmbAHName.BackColor = &HFFFFC0
End Sub

Private Sub cmbAHName_LostFocus()
    cmbAHName.BackColor = vbWhite
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    CmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdNew.Caption = "&New"
    cmdEdit.Caption = "&Edit"
    cmdPrint.Enabled = True
    cmdPreview.Enabled = True
'    CmdDelete.Enabled = True
    cmdFind.Enabled = True
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdPost.Enabled = True
'    cmdChange.Enabled = True
    txtVID.Enabled = False
    Call allClear
    Call alldisable
    If Not rsfactory.EOF Then FindRecord
End Sub
'
Private Sub cmdNew_Click()
    Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        CmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdFind.Enabled = False
        cmdPreview.Enabled = False
        cmdPost.Enabled = False
        cmdPreview.Enabled = False
        txtUName.text = frmLogIn.Txtuserid.text
        txtCPost.text = "Not Posted"
        Call allClear

If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(VID),0) as InvNo from Voucher"
            rs.Open str, cn, adOpenStatic, adLockReadOnly
                txtVID.text = Val(rs!InvNo) + 1

        Call allenable
            cmbAHName.SetFocus
        
    ElseIf cmdNew.Caption = "&Save" Then
    Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtVID.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                CmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdPrint.Enabled = True
                cmdFind.Enabled = True
                cmdPreview.Enabled = True
                cmdPost.Enabled = True
                Call alldisable
                
                s = txtVID
                rsfactory.Requery
                rsfactory.MoveFirst
                rsfactory.Find "VID='" & parseQuotes(s) & "'"
                
                FindRecord

            End If
        End If
    End If
 
ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select
End Sub

Private Sub cmdEdit_Click()


If cmdEdit.Caption = "&Edit" Then
     strMood = "U"
    If txtCPost.text = "Not Posted" Then
        cmdNew.Enabled = False
        Call allenable
        cmbAHName.SetFocus
        cmdNew.Enabled = False
        Call allenable
'        cmbAHName.SetFocus
        cmdEdit.Caption = "&Update"
        CmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdPreview.Enabled = False
        cmdFind.Enabled = False
        cmdPost.Enabled = False
        cmdPrint.Enabled = False
        txtVID.Enabled = False
        
      End If

  ElseIf cmdEdit.Caption = "&Update" Then
'  Call Calculation
'    Call duplicate
'    If txtCPost.text = "Not Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                CmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdPreview.Enabled = True
                cmdFind.Enabled = True
                cmdPost.Enabled = True
                cmdPreview.Enabled = True
                cmdPrint.Enabled = True
                Call alldisable
                rsfactory.Requery

                Dim s As String
                s = cmbAHName
                rsfactory.Find "AHead='" & parseQuotes(s) & "'"
            End If
        End If
'   Call Calculation
'     End If
    
End If
'If txtUName.text <> "Admin" Then
'                MsgBox "If you want to Change the Information, you should contact to Managing Director.....", vbCritical
'            Exit Sub
'    If cmdEdit.Caption = "&Edit" Then
'
'        cmdNew.Enabled = False
'        Call allenable
'        cmbAHName.SetFocus
'        cmdEdit.Caption = "&Update"
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
'        cmdPreview.Enabled = False
'        cmdFind.Enabled = False
'        cmdPost.Enabled = False
'        cmdPrint.Enabled = False
'
'ElseIf cmdEdit.Caption = "&Update" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdEdit.Caption = "&Edit"
'                cmdNew.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                cmdPreview.Enabled = True
'                cmdFind.Enabled = True
'                cmdPost.Enabled = True
'                cmdPreview.Enabled = True
'                cmdPrint.Enabled = True
'                Call alldisable
'                rsfactory.Requery
'
'                Dim s As String
'                s = cmbAHName
'                rsfactory.Find "AHead='" & parseQuotes(s) & "'"
''                Call search
''                Call countrysearch
'                FindRecord
'
'            End If
'        End If
'    End If
'    End If
End Sub

Private Sub cmdFind_Click()
'    frmVoucherSearch.Show vbModal
    cmdFind.Enabled = True
    CmdCancel.Enabled = True
End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
'          MsgBox "end of file"
       cmdPrevious.Enabled = False
 Else
      cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtVID = Adodc1.Recordset!VID
        VDate.value = Adodc1.Recordset!VDate
        cmbAHName = Adodc1.Recordset!AHead
        txtDebit = Adodc1.Recordset!Debit
        txtCredit = Adodc1.Recordset!Credit
        cmbAMode = Adodc1.Recordset!AMode
        txtDescription = Adodc1.Recordset!Description
        txtTime = Adodc1.Recordset!StrTIME
        txtUName = Adodc1.Recordset!UName
        txtCPost = Adodc1.Recordset!Posted

End If
End Sub

Private Sub cmdPrint_Click()
Dim s As String
If cmdPrint.Caption = "&Print" Then
cmdPrint.Caption = "&Printing"
        If IsValidRecord Then
            If rcupdate Then
'                cmdPrint.Caption = "&Printing"
                cmdEdit.Enabled = True
                CmdCancel.Enabled = True
                cmdClose.Enabled = True
                cmdPost.Enabled = True
                cmdFind.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
                txtVID.Enabled = False
                Call alldisable
'                txtWords = InWords(txtNPayable.text)

            End If
        End If
    End If

Tracer = 1
Screen.MousePointer = vbHourglass
If txtCPost.text = "Posted" Then
If txtCredit.text = "0" Then
Call Debit
Else
Call Credit
End If
Screen.MousePointer = vbDefault

cmdPrint.Caption = "&Print"

End If
End Sub


Private Sub Timer1_Timer()
    txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
End Sub

Private Sub Form_Load()

    Call Connect
       ModFunction.StartUpPosition Me

       txtUName.text = frmLogIn.Txtuserid.text
       Call alldisable
       Call AccountHead
'       Call UndoPostVisible

       txtCPost.text = "Not Posted"
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from Voucher", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If

    If Not rsfactory.EOF Then FindRecord

    txtVID.Enabled = False
    VDate.value = Date
    txtTime.text = Time
    
    cmbAMode.AddItem "Expenditure"
    cmbAMode.AddItem "Income"
    cmbAMode.AddItem "Fixed Assets"
'    cmbAMode.AddItem "Bank"
    
    
    Adodc1.connectionstring = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  Adodc1.CommandType = adCmdTable
  Adodc1.RecordSource = "Voucher"

  Adodc1.Refresh

End Sub

Private Sub cmbAHName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbAHName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If

End Sub
'
Private Sub AccountHead()

Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT AHName FROM AccountsHead ORDER BY AHName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbAHName.AddItem rsTemp2("AHName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

Private Sub allenable()
    txtDebit.Enabled = True
    txtCredit.Enabled = True
    cmbAHName.Enabled = True
    cmbAMode.Enabled = True
    VDate.Enabled = True
    txtDescription.Enabled = True
'    Check1.Enabled = True
End Sub

Private Sub alldisable()
    txtVID.Enabled = False
    txtDebit.Enabled = False
    cmbAHName.Enabled = False
    txtCredit.Enabled = False
    cmbAMode.Enabled = False
    txtCPost.Enabled = False
    txtUName.Enabled = False
    VDate.Enabled = False
    txtDescription.Enabled = False
End Sub

Private Sub allClear()
    cmbAHName.text = ""
    txtDebit.text = "0"
    txtCredit.text = "0"
    txtDescription.text = ""
    cmbAMode.text = ""
    VDate.value = Date
End Sub

Private Function rcupdate() As Boolean

On Error Resume Next

Dim ipost
Dim iprint

cn.BeginTrans

    If cmdNew.Caption = "&Save" Then

    cn.Execute "INSERT INTO Voucher(VID,VDate,AHead,Debit,Credit,AMode,Description,strTime, " & _
                   " Posted,UName) " & _
                   " VALUES ('" & txtVID & "','" & Format(VDate, "dd-mmm-yyyy") & "','" & parseQuotes(cmbAHName) & "'," & _
                   " " & Val(txtDebit.text) & "," & Val(txtCredit.text) & "," & _
                   " '" & parseQuotes(cmbAMode) & "','" & parseQuotes(txtDescription) & "','" & txtTime.text & "','" & txtCPost.text & "','" & txtUName.text & "') "


          rcupdate = True
          cn.CommitTrans
          MsgBox "Record Added", vbInformation, "Confirmation"

    ElseIf (cmdEdit.Caption = "&Update") Then

    cn.Execute "Update Voucher SET VDate='" & Format(VDate, "dd-mmm-yyyy") & "',AHead='" & parseQuotes(cmbAHName) & "', " & _
                   " Debit=" & Val(txtDebit.text) & ",Credit=" & Val(txtCredit.text) & ",AMode='" & parseQuotes(cmbAMode) & "', " & _
                   " strTime='" & (txtTime.text) & "',Posted='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "' WHERE VID = '" & txtVID & "'"

        rcupdate = True
        cn.CommitTrans
        MsgBox "Record Updated", vbInformation, "Confirmation"

'    End If

'----------------------------------------------Printing Start--------------------------
  ElseIf cmdPrint.Caption = "&Printing" Then

    txtCPost.text = "Posted"

'    iprint = MsgBox("Do you want to Print this Money Receipt?", vbYesNo)

    If iprint = vbYes Then

         cn.Execute "Update Voucher SET VDate='" & Format(VDate, "dd-mmm-yyyy") & "',AHead='" & parseQuotes(cmbAHName) & "', " & _
                   " Debit=" & Val(txtDebit.text) & ",Credit=" & Val(txtCredit.text) & ",AMode='" & parseQuotes(cmbAMode) & "', " & _
                   "strTime='" & (txtTime.text) & "',Posted='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "' WHERE VID = '" & txtVID & "'"


        rcupdate = True
'        cn.CommitTrans
'
        End If
'----------------------------------Printing End---------------------------

'----------------------------------Posted Start--------------------------
    ElseIf cmdPost.Caption = "&Posted" Then

     txtCPost.text = "Posted"

     ipost = MsgBox("Do you want to Post this bill?", vbYesNo)

           If ipost = vbYes Then

     cn.Execute "Update Voucher SET VDate='" & Format(VDate, "dd-mmm-yyyy") & "',AHead='" & parseQuotes(cmbAHName) & "', " & _
                   " Debit=" & Val(txtDebit.text) & ",Credit=" & Val(txtCredit.text) & ",AMode='" & parseQuotes(cmbAMode) & "', " & _
                   "strTime='" & (txtTime.text) & "',Posted='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "' WHERE VID = '" & txtVID & "'"

       rcupdate = True
       cn.CommitTrans
       MsgBox "Record Posted Successfully", vbInformation, "Confirmation"

    End If
        End If
            Exit Function
            
End Function
Public Sub FindRecord()

If Not rsfactory.EOF Then
        txtVID = rsfactory("VID")
        cmbAHName = rsfactory("AHead")
        VDate = rsfactory("VDate")
        txtDebit = rsfactory("Debit") & ""
        txtCredit = rsfactory("Credit") & ""
        cmbAMode = rsfactory("AMode") & ""
        txtDescription = rsfactory("Description")
        txtTime = rsfactory("StrTime") & ""
        txtUName = IIf(IsNull(rsfactory("UName")), "", rsfactory("UName"))
        txtCPost = IIf(IsNull(rsfactory("Posted")), "", rsfactory("Posted"))
    End If

End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If (cmbAHName.text = "") Then
       MsgBox "Enter Accounts Head Name"
       cmbAHName.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtDebit.text = "") Then
      MsgBox "Enter Debit Amount"
      txtDebit.SetFocus
      IsValidRecord = False
      Exit Function
    End If

    If (txtCredit.text = "") Then
      MsgBox "Enter Credit Amount"
      txtCredit.SetFocus
      IsValidRecord = False
      Exit Function
    End If


    End Function
'.............................................................................

Public Sub Debit()
'On Error GoTo ErrorHan
Dim strPath         As String
Dim rsFactProf      As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\Debit Voucher.rpt"

    Set objReportApp = CreateObject("Crystal.CRPE.Application")
    Set objReport = objReportApp.OpenReport(strPath)
    Set objReportDatabase = objReport.Database
    Set objReportDatabaseTables = objReportDatabase.Tables
    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = objReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select Voucher.VID,Voucher.VDate,Voucher.AHead, " & _
             "  " & _
             "Voucher.Description,Voucher.Debit,Voucher.Credit,Voucher.UName " & _
             "from Voucher where " & _
             "Voucher.VID='" & Me.txtVID & "'"

    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    objReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True

'      Set oReportFormulaFieldDefinations = oReport.FormulaFields
'      Set oReportFF = oReportFormulaFieldDefinations.Item(1)
'      oReportFF.text = "'Factory Information'"

objReport.DiscardSavedData
objReport.Preview "Money Receipt Infromation of '" & cmbAHName.text & "'", , , , , 16777216 Or 524288 Or 65536

End Sub

Public Sub PopulateCnf(StrID As String)
    rsfactory.MoveFirst
    rsfactory.Find "CID=" & parseQuotes(StrID)
    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub

Public Sub Credit()
'On Error GoTo ErrorHan
Dim strPath         As String
Dim rsFactProf      As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\Credit Voucher.rpt"

    Set objReportApp = CreateObject("Crystal.CRPE.Application")
    Set objReport = objReportApp.OpenReport(strPath)
    Set objReportDatabase = objReport.Database
    Set objReportDatabaseTables = objReportDatabase.Tables
    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = objReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select Voucher.VID,Voucher.VDate,Voucher.AHead, " & _
             "  " & _
             "Voucher.Description,Voucher.Debit,Voucher.Credit " & _
             "from Voucher where " & _
             "Voucher.VID='" & Me.txtVID & "'"

    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    objReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True

'      Set oReportFormulaFieldDefinations = oReport.FormulaFields
'      Set oReportFF = oReportFormulaFieldDefinations.Item(1)
'      oReportFF.text = "'Factory Information'"

objReport.DiscardSavedData
objReport.Preview "Voucher Infromation of '" & cmbAHName.text & "'", , , , , 16777216 Or 524288 Or 65536

End Sub

