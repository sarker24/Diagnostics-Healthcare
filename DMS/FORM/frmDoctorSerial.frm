VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDoctorSerial 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Doctor's Serial Entry Informations"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   Icon            =   "frmDoctorSerial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2775
      Left            =   0
      TabIndex        =   30
      Top             =   3240
      Width           =   10455
      Begin VB.TextBox txtDescription 
         Height          =   2265
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Text            =   " "
         Top             =   360
         Width           =   10215
      End
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
      Picture         =   "frmDoctorSerial.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
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
      Picture         =   "frmDoctorSerial.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6600
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
      Picture         =   "frmDoctorSerial.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6600
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
      Picture         =   "frmDoctorSerial.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   1095
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
      Picture         =   "frmDoctorSerial.frx":2334
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      UseMaskColor    =   -1  'True
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
      Picture         =   "frmDoctorSerial.frx":2BFE
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6600
      Width           =   990
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
      Picture         =   "frmDoctorSerial.frx":34C8
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6600
      Width           =   1095
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
      MouseIcon       =   "frmDoctorSerial.frx":3D92
      Picture         =   "frmDoctorSerial.frx":447C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Find Next"
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Find Last"
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Find Previous"
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Find First"
      Top             =   6120
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   10080
      Top             =   6960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   3135
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10455
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmDoctorSerial.frx":4B66
         Left            =   1680
         List            =   "frmDoctorSerial.frx":4B68
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cmbMExecutive 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2640
         Width           =   5895
      End
      Begin VB.TextBox txtPhoneNo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   390
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtAge 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   390
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtPatientName 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   6855
      End
      Begin VB.TextBox txtPID 
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
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtAmount 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   390
         Left            =   6480
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2040
         Width           =   1815
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
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1455
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
         Height          =   390
         Left            =   8760
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "frmDoctorSerial.frx":4B6A
         Top             =   1320
         Width           =   1455
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
         Left            =   8760
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbDName 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2040
         Width           =   5895
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
         Left            =   6480
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
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
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   63111171
         CurrentDate     =   39739
      End
      Begin VB.Label lblSex 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   36
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblMExecutive 
         BackColor       =   &H00C0B4A9&
         Caption         =   "M Executive Name"
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
         TabIndex        =   35
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Label lblPhoneNo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Phone No"
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
         Left            =   3480
         TabIndex        =   34
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lblConsultantFee 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Consultant Fee"
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
         Left            =   6480
         TabIndex        =   33
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Age"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblPatientName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Patient Name"
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
         Left            =   1680
         TabIndex        =   31
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblPID 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Patient ID"
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
         TabIndex        =   20
         Top             =   360
         Width           =   1455
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
         Left            =   6480
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblDoctorName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Doctor's Name"
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
         TabIndex        =   18
         Top             =   1800
         Width           =   3735
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
         Left            =   8760
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7920
      Top             =   6120
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
End
Attribute VB_Name = "frmDoctorSerial"
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




'Private Sub cmdFirst_Click()
'Adodc1.Recordset.MoveFirst
'If Adodc1.Recordset.EOF = True Then
'       cmdFirst.Enabled = False
' Else
'       cmdFirst.Enabled = True
'       cmdNext.Enabled = True
'       cmdLast.Enabled = True
'       cmdPrevious.Enabled = True
'
'        txtPID = Adodc1.Recordset!PID
'        VDate.value = Adodc1.Recordset!VDate
'        cmbDName = Adodc1.Recordset!Doctor
'        txtPatientName = Adodc1.Recordset!PatientName
'        txtAge = Adodc1.Recordset!age
'        txtAmount = Adodc1.Recordset!amount
'        txtPhoneNo = Adodc1.Recordset!PhoneNo
'        txtDescription = Adodc1.Recordset!Description
'        txtTime = Adodc1.Recordset!StrTIME
'        txtUName = Adodc1.Recordset!UName
'        txtCPost = Adodc1.Recordset!Posted
'End If
'
'
''End If
'End Sub
'
'Private Sub cmdLast_Click()
'Adodc1.Recordset.MoveLast
'If Adodc1.Recordset.EOF = True Then
''          MsgBox "end of file"
'       cmdLast.Enabled = False
' Else
'       cmdFirst.Enabled = True
'       cmdNext.Enabled = True
'       cmdLast.Enabled = True
'       cmdPrevious.Enabled = True
'
'       txtPID = Adodc1.Recordset!PID
'        VDate.value = Adodc1.Recordset!VDate
'        cmbDName = Adodc1.Recordset!Doctor
'        txtPatientName = Adodc1.Recordset!PatientName
'        txtAge = Adodc1.Recordset!age
'        txtAmount = Adodc1.Recordset!amount
'        txtPhoneNo = Adodc1.Recordset!PhoneNo
'        txtDescription = Adodc1.Recordset!Description
'        txtTime = Adodc1.Recordset!StrTIME
'        txtUName = Adodc1.Recordset!UName
'        txtCPost = Adodc1.Recordset!Posted
'
'End If
'End Sub
'
'Private Sub cmdNext_Click()
'Adodc1.Recordset.MoveNext
'If Adodc1.Recordset.EOF = True Then
''          MsgBox "end of file"
'       cmdNext.Enabled = False
' Else
'       cmdFirst.Enabled = True
'       cmdNext.Enabled = True
'       cmdLast.Enabled = True
'       cmdPrevious.Enabled = True
'
'       txtPID = Adodc1.Recordset!PID
'        VDate.value = Adodc1.Recordset!VDate
'        cmbDName = Adodc1.Recordset!Doctor
'        txtPatientName = Adodc1.Recordset!PatientName
'        txtAge = Adodc1.Recordset!age
'        txtAmount = Adodc1.Recordset!amount
'        txtPhoneNo = Adodc1.Recordset!PhoneNo
'        txtDescription = Adodc1.Recordset!Description
'        txtTime = Adodc1.Recordset!StrTIME
'        txtUName = Adodc1.Recordset!UName
'        txtCPost = Adodc1.Recordset!Posted
'        End If
'End Sub
'
'Private Sub cmdPost_Click()
'Dim s As String
'
'cmdPost.Caption = "&Posted"
'
'If cmdPost.Caption = "&Posted" Then
'     If txtCPost.text = "Not Posted" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                 cmdNew.Caption = "&New"
'                 cmdEdit.Enabled = True
'                 cmdCancel.Enabled = False
'                 cmdClose.Enabled = True
'                 cmdFind.Enabled = False
'                 cmdPreview.Enabled = True
'                 cmdPrint.Enabled = True
'                 txtPID.Enabled = False
'                 Call alldisable
'           End If
'        End If
'      End If
'Else
' End If
'cmdPost.Caption = "&Posted"
'
'End Sub
'
Private Sub CmdPreview_Click()
'    Call printReport
 Tracer = 0
''    Call printReport
'If txtCPost.text = "Posted" Then
'  If txtAmount.text = "0" Then
'   Call Debit
'
'   Else
'   Call Credit
'   End If
'   End If

End Sub


Private Sub cmbDName_GotFocus()
cmbDName.BackColor = &HFFFFC0
End Sub

Private Sub cmbDName_LostFocus()
    cmbDName.BackColor = vbWhite
End Sub

'Private Sub cmdClose_Click()
'    Unload Me
'End Sub
'Private Sub cmdCancel_Click()
'    cmdCancel.Enabled = False
'    cmdNew.Enabled = True
'    cmdNew.Caption = "&New"
'    cmdEdit.Caption = "&Edit"
'    cmdPrint.Enabled = True
'    cmdPreview.Enabled = True
''    CmdDelete.Enabled = True
'    cmdFind.Enabled = True
'    cmdClose.Enabled = True
'    cmdEdit.Enabled = True
'    cmdPost.Enabled = True
''    cmdChange.Enabled = True
'    txtPID.Enabled = False
'    Call allClear
'    Call alldisable
'    If Not rsfactory.EOF Then FindRecord
'End Sub
''
'Private Sub cmdNew_Click()
'    Set rs = New ADODB.Recordset
'    If cmdNew.Caption = "&New" Then
'        cmdNew.Caption = "&Save"
'        cmdEdit.Enabled = False
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
'        cmdFind.Enabled = False
'        cmdPreview.Enabled = False
'        cmdPost.Enabled = False
'        cmdPreview.Enabled = False
'        txtUName.text = frmLogIn.Txtuserid.text
'        txtCPost.text = "Not Posted"
'        Call allClear
'
'If rs.State <> 0 Then rs.Close
'           str = "Select ISNULL(max(PID),0) as InvNo from Doctor_Serial"
'            rs.Open str, cn, adOpenStatic, adLockReadOnly
'                txtPID.text = Val(rs!InvNo) + 1
'
'        Call allenable
'            txtPatientName.SetFocus
'
'    ElseIf cmdNew.Caption = "&Save" Then
'    Dim s As String
'        If IsValidRecord Then
'            If rcupdate Then
'                txtPID.Enabled = False
'                cmdNew.Caption = "&New"
'                cmdEdit.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                cmdPrint.Enabled = True
'                cmdFind.Enabled = True
'                cmdPreview.Enabled = True
'                cmdPost.Enabled = True
'                Call alldisable
'
''                s = txtPID
''                rsfactory.Requery
''                rsfactory.MoveFirst
''                rsfactory.Find "PID='" & parseQuotes(s) & "'"
'
'                FindRecord
'
'            End If
'        End If
'    End If
'
'ProcError:
'    Select Case Err.Number
'    Case 0:
'    Case Else
'        MsgBox Err.Description
'    End Select
'End Sub
'
'Private Sub cmdEdit_Click()
'
'
'If cmdEdit.Caption = "&Edit" Then
'     strMood = "U"
'    If txtCPost.text = "Not Posted" Then
'        cmdNew.Enabled = False
'        Call allenable
'        cmbDName.SetFocus
'        cmdNew.Enabled = False
'        Call allenable
''        cmbDName.SetFocus
'        cmdEdit.Caption = "&Update"
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
'        cmdPreview.Enabled = False
'        cmdFind.Enabled = False
'        cmdPost.Enabled = False
'        cmdPrint.Enabled = False
'        txtPID.Enabled = False
'
'      End If
'
'  ElseIf cmdEdit.Caption = "&Update" Then
''  Call Calculation
''    Call duplicate
''    If txtCPost.text = "Not Posted" Then
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
'                s = cmbDName
'                rsfactory.Find "AHead='" & parseQuotes(s) & "'"
'            End If
'        End If
''   Call Calculation
''     End If
'
'End If
''If txtUName.text <> "Admin" Then
''                MsgBox "If you want to Change the Information, you should contact to Managing Director.....", vbCritical
''            Exit Sub
''    If cmdEdit.Caption = "&Edit" Then
''
''        cmdNew.Enabled = False
''        Call allenable
''        cmbDName.SetFocus
''        cmdEdit.Caption = "&Update"
''        cmdCancel.Enabled = True
''        cmdClose.Enabled = False
''        cmdPreview.Enabled = False
''        cmdFind.Enabled = False
''        cmdPost.Enabled = False
''        cmdPrint.Enabled = False
''
''ElseIf cmdEdit.Caption = "&Update" Then
''        If IsValidRecord Then
''            If rcupdate Then
''                cmdEdit.Caption = "&Edit"
''                cmdNew.Enabled = True
''                cmdCancel.Enabled = False
''                cmdClose.Enabled = True
''                cmdPreview.Enabled = True
''                cmdFind.Enabled = True
''                cmdPost.Enabled = True
''                cmdPreview.Enabled = True
''                cmdPrint.Enabled = True
''                Call alldisable
''                rsfactory.Requery
''
''                Dim s As String
''                s = cmbDName
''                rsfactory.Find "AHead='" & parseQuotes(s) & "'"
'''                Call search
'''                Call countrysearch
''                FindRecord
''
''            End If
''        End If
''    End If
''    End If
'End Sub
'
'Private Sub cmdFind_Click()
''    frmDoctor_SerialSearch.Show vbModal
'    cmdFind.Enabled = True
'    cmdCancel.Enabled = True
'End Sub
'
'Private Sub cmdPrevious_Click()
'Adodc1.Recordset.MovePrevious
'If Adodc1.Recordset.BOF = True Then
''          MsgBox "end of file"
'       cmdPrevious.Enabled = False
' Else
'      cmdFirst.Enabled = True
'       cmdNext.Enabled = True
'       cmdLast.Enabled = True
'       cmdPrevious.Enabled = True
'
'       txtPID = Adodc1.Recordset!PID
'        VDate.value = Adodc1.Recordset!VDate
'        cmbDName = Adodc1.Recordset!Doctor
'        txtPatientName = Adodc1.Recordset!PatientName
'        txtAge = Adodc1.Recordset!age
'        txtAmount = Adodc1.Recordset!amount
'        txtPhoneNo = Adodc1.Recordset!PhoneNo
'        txtDescription = Adodc1.Recordset!Description
'        txtTime = Adodc1.Recordset!StrTIME
'        txtUName = Adodc1.Recordset!UName
'        txtCPost = Adodc1.Recordset!Posted
'
'End If
'End Sub
'
'Private Sub cmdPrint_Click()
'Dim s As String
'If cmdPrint.Caption = "&Print" Then
'cmdPrint.Caption = "&Printing"
'        If IsValidRecord Then
'            If rcupdate Then
''                cmdPrint.Caption = "&Printing"
'                cmdEdit.Enabled = True
'                cmdCancel.Enabled = True
'                cmdClose.Enabled = True
'                cmdPost.Enabled = True
'                cmdFind.Enabled = True
'                cmdPrint.Enabled = True
'                cmdPreview.Enabled = True
'                txtPID.Enabled = False
'                Call alldisable
''                txtWords = InWords(txtNPayable.text)
'
'            End If
'        End If
'    End If
'
'Tracer = 1
'Screen.MousePointer = vbHourglass
'If txtCPost.text = "Posted" Then
'If txtCredit.text = "0" Then
'Call Debit
'Else
'Call Credit
'End If
'Screen.MousePointer = vbDefault
'
'cmdPrint.Caption = "&Print"
'
'End If
'End Sub
'
'
'Private Sub Timer1_Timer()
'    txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
'End Sub
'
'Private Sub Form_Load()
'
'    Call Connect
'       ModFunction.StartUpPosition Me
'
'       txtUName.text = frmLogIn.Txtuserid.text
'       Call alldisable
'       Call AccountHead
''       Call UndoPostVisible
'
'       txtCPost.text = "Not Posted"
'    Set rsfactory = New ADODB.Recordset
'    rsfactory.Open "select * from Doctor_Serial", cn, adOpenStatic, adLockReadOnly
'    Call alldisable
'   If rsfactory.RecordCount > 0 Then
'        bRecordExists = True
'    Else
'        bRecordExists = False
'    End If
'
'    If Not rsfactory.EOF Then FindRecord
'
'    txtPID.Enabled = False
'    VDate.value = Date
'    txtTime.text = Time
'
''    cmbAMode.AddItem "Expenditure"
''    cmbAMode.AddItem "Income"
''    cmbAMode.AddItem "Tution Fee"
'
'
'    Adodc1.connectionstring = "Driver={SQL Server};" & _
'           "Server=" & sServerName & ";" & _
'           "Database=" & SDatabaseName & ";" & _
'           "Trusted_Connection=yes"
'
'  Adodc1.CommandType = adCmdTable
'  Adodc1.RecordSource = "Doctor_Serial"
'
'  Adodc1.Refresh
'
'End Sub
'
'Private Sub cmbDName_KeyPress(KeyAscii As Integer)
'   KeyAscii = AutoMatchCBBox(cmbDName, KeyAscii)
'   If KeyAscii = 13 Then
'       SendKeys Chr(9)
'    End If
'
'End Sub
''
'Private Sub AccountHead()
'
'Dim rsTemp2 As New ADODB.Recordset
'
'     rsTemp2.Open ("SELECT DISTINCT doc_name FROM Doctor_Info ORDER BY doc_name ASC"), cn, adOpenStatic
'
'    While Not rsTemp2.EOF
'        cmbDName.AddItem rsTemp2("doc_name")
'        rsTemp2.MoveNext
'    Wend
'    rsTemp2.Close
'
'End Sub
'
'Private Sub allenable()
'    txtPID.Enabled = True
'    txtPatientName.Enabled = True
'    txtAge.Enabled = True
'    cmbDName.Enabled = True
'    txtPhoneNo.Enabled = True
'    VDate.Enabled = True
'    txtDescription.Enabled = True
'    txtAmount.Enabled = True
'End Sub
'
'Private Sub alldisable()
'    txtPID.Enabled = False
'    txtAge.Enabled = False
'    txtPatientName.Enabled = False
'    txtAmount.Enabled = False
'    txtPhoneNo.Enabled = False
'    cmbDName.Enabled = False
'    txtCPost.Enabled = False
'    txtUName.Enabled = False
'    VDate.Enabled = False
'    txtDescription.Enabled = False
'End Sub
'
'Private Sub allClear()
'    cmbDName.text = ""
'    txtPatientName.text = ""
'    txtAge.text = ""
'    txtDescription.text = ""
'    txtAmount.text = ""
'    txtPhoneNo.text = ""
'    VDate.value = Date
'End Sub
'
'Private Function rcupdate() As Boolean
'
'On Error Resume Next
'
'Dim ipost
'Dim iprint
'
'cn.BeginTrans
'
'    If cmdNew.Caption = "&Save" Then
'
'    cn.Execute "INSERT INTO Doctor_Serial(PID,VDate,PatientName,Doctor,Age,PhoneNo,Amount,Description,strTime, " & _
'                   " Posted,UName) " & _
'                   " VALUES ('" & txtPID & "','" & Format(VDate, "dd-MM-yyyy") & "','" & parseQuotes(txtPatientName) & "'," & _
'                   " '" & parseQuotes(cmbDName) & "', " & _
'                   " " & Val(txtPhoneNo.text) & "," & Val(txtAge.text) & "," & Val(txtAmount.text) & "," & _
'                   " '" & parseQuotes(txtPatientName) & "','" & parseQuotes(txtDescription) & "','" & txtTime.text & "','" & txtCPost.text & "','" & txtUName.text & "') "
'
'
'          rcupdate = True
'          cn.CommitTrans
'          MsgBox "Record Added", vbInformation, "Confirmation"
'
'    ElseIf (cmdEdit.Caption = "&Update") Then
'
'    cn.Execute "Update Doctor_Serial SET VDate='" & Format(VDate, "dd-mmm-yyyy") & "',PatientName='" & parseQuotes(txtPatientName.text) & "',Doctor='" & parseQuotes(cmbDName) & "', " & _
'                   " PhoneNo=" & Val(txtPhoneNo.text) & ",Age=" & Val(txtAge.text) & ",Amount=" & Val(txtAmount.text) & "', " & _
'                   " strTime='" & (txtTime.text) & "',Posted='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "' WHERE PID = '" & txtPID & "'"
'
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record Updated", vbInformation, "Confirmation"
'
''    End If
'
''----------------------------------------------Printing Start--------------------------
'  ElseIf cmdPrint.Caption = "&Printing" Then
'
'    txtCPost.text = "Posted"
'
''    iprint = MsgBox("Do you want to Print this Money Receipt?", vbYesNo)
'
'    If iprint = vbYes Then
'
'    cn.Execute "Update Doctor_Serial SET VDate='" & Format(VDate, "dd-mmm-yyyy") & "',PatientName='" & parseQuotes(txtPatientName.text) & "',Doctor='" & parseQuotes(cmbDName) & "', " & _
'                   " PhoneNo=" & Val(txtPhoneNo.text) & ",Age=" & Val(txtAge.text) & ",Amount=" & Val(txtAmount.text) & "', " & _
'                   " strTime='" & (txtTime.text) & "',Posted='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "' WHERE PID = '" & txtPID & "'"
'
'        rcupdate = True
''        cn.CommitTrans
''
'        End If
''----------------------------------Printing End---------------------------
'
''----------------------------------Posted Start--------------------------
'    ElseIf cmdPost.Caption = "&Posted" Then
'
'     txtCPost.text = "Posted"
'
'     ipost = MsgBox("Do you want to Post this bill?", vbYesNo)
'
'           If ipost = vbYes Then
'
'     cn.Execute "Update Doctor_Serial SET VDate='" & Format(VDate, "dd-mmm-yyyy") & "',PatientName='" & parseQuotes(txtPatientName.text) & "',Doctor='" & parseQuotes(cmbDName) & "', " & _
'                   " PhoneNo=" & Val(txtPhoneNo.text) & ",Age=" & Val(txtAge.text) & ",Amount=" & Val(txtAmount.text) & "', " & _
'                   " strTime='" & (txtTime.text) & "',Posted='" & parseQuotes(txtCPost.text) & "',UName='" & txtUName.text & "' WHERE PID = '" & txtPID & "'"
'
'
'       rcupdate = True
'       cn.CommitTrans
'       MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
'
'    End If
'        End If
'            Exit Function
'
'End Function
'Public Sub FindRecord()
'
'If Not rsfactory.EOF Then
'        txtPID = rsfactory("PID")
'        txtPatientName = rsfactory("PatientName") & ""
'        txtAge = rsfactory("Age") & ""
'        txtPhoneNo = rsfactory("PhoneNo")
'        cmbDName = rsfactory("Doctor")
'        VDate = rsfactory("VDate")
'        txtAmount = rsfactory("Amount") & ""
'        txtDescription = rsfactory("Description")
'        txtTime = rsfactory("StrTime") & ""
'        txtUName = IIf(IsNull(rsfactory("UName")), "", rsfactory("UName"))
'        txtCPost = IIf(IsNull(rsfactory("Posted")), "", rsfactory("Posted"))
'    End If
'
'End Sub
'
'Private Function IsValidRecord() As Boolean
'    IsValidRecord = True
'    If (txtPatientName.text = "") Then
'       MsgBox "Enter Patient Name."
'       txtPatientName.SetFocus
'       IsValidRecord = False
'       Exit Function
'    End If
'
'    If (txtAge.text = "") Then
'      MsgBox "Enter Patient Age."
'      txtAge.SetFocus
'      IsValidRecord = False
'      Exit Function
'    End If
'
'    If (txtPhoneNo.text = "") Then
'      MsgBox "Enter Patient Phone No"
'      txtPhoneNo.SetFocus
'      IsValidRecord = False
'      Exit Function
'    End If
'
'    If (cmbDName.text = "") Then
'      MsgBox "Enter Doctor Name."
'      cmbDName.SetFocus
'      IsValidRecord = False
'      Exit Function
'    End If
'
'    If (txtAmount.text = "") Then
'      MsgBox "Enter Doctor Fee."
'      txtAmount.SetFocus
'      IsValidRecord = False
'      Exit Function
'    End If
'
'
'    End Function
''.............................................................................
'
'Public Sub Debit()
''On Error GoTo ErrorHan
'Dim strPath         As String
'Dim rsFactProf      As ADODB.Recordset
'Dim strSQL          As String
'
'
'    strPath = App.Path + "\reports\Debit Doctor_Serial.rpt"
'
'    Set objReportApp = CreateObject("Crystal.CRPE.Application")
'    Set objReport = objReportApp.OpenReport(strPath)
'    Set objReportDatabase = objReport.Database
'    Set objReportDatabaseTables = objReportDatabase.Tables
'    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
'    Set ObjPrinterSetting = objReport.PrintWindowOptions
'
'
'    Set rsFactProf = New ADODB.Recordset
'If rsFactProf.State <> 0 Then rsFactProf.Close
'
'    strSQL = "select Doctor_Serial.PID,Doctor_Serial.VDate,Doctor_Serial.AHead, " & _
'             "  " & _
'             "Doctor_Serial.Description,Doctor_Serial.Debit,Doctor_Serial.Credit,Doctor_Serial.UName " & _
'             "from Doctor_Serial where " & _
'             "Doctor_Serial.PID='" & Me.txtPID & "'"
'
'    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly
'
'    objReportDatabaseTable.SetPrivateData 3, rsFactProf
'
'ObjPrinterSetting.HasPrintSetupButton = True
'ObjPrinterSetting.HasRefreshButton = True
'ObjPrinterSetting.HasSearchButton = True
'ObjPrinterSetting.HasZoomControl = True
'
''      Set oReportFormulaFieldDefinations = oReport.FormulaFields
''      Set oReportFF = oReportFormulaFieldDefinations.Item(1)
''      oReportFF.text = "'Factory Information'"
'
'objReport.DiscardSavedData
'objReport.Preview "Money Receipt Infromation of '" & cmbDName.text & "'", , , , , 16777216 Or 524288 Or 65536
'
'End Sub
'
'Public Sub PopulateCnf(StrID As String)
'    rsfactory.MoveFirst
'    rsfactory.Find "CID=" & parseQuotes(StrID)
'    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord
'
'End Sub
'
'Public Sub Credit()
''On Error GoTo ErrorHan
'Dim strPath         As String
'Dim rsFactProf      As ADODB.Recordset
'Dim strSQL          As String
'
'
'    strPath = App.Path + "\reports\Credit Doctor_Serial.rpt"
'
'    Set objReportApp = CreateObject("Crystal.CRPE.Application")
'    Set objReport = objReportApp.OpenReport(strPath)
'    Set objReportDatabase = objReport.Database
'    Set objReportDatabaseTables = objReportDatabase.Tables
'    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
'    Set ObjPrinterSetting = objReport.PrintWindowOptions
'
'
'    Set rsFactProf = New ADODB.Recordset
'If rsFactProf.State <> 0 Then rsFactProf.Close
'
'    strSQL = "select Doctor_Serial.PID,Doctor_Serial.VDate,Doctor_Serial.AHead, " & _
'             "  " & _
'             "Doctor_Serial.Description,Doctor_Serial.Debit,Doctor_Serial.Credit " & _
'             "from Doctor_Serial where " & _
'             "Doctor_Serial.PID='" & Me.txtPID & "'"
'
'    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly
'
'    objReportDatabaseTable.SetPrivateData 3, rsFactProf
'
'ObjPrinterSetting.HasPrintSetupButton = True
'ObjPrinterSetting.HasRefreshButton = True
'ObjPrinterSetting.HasSearchButton = True
'ObjPrinterSetting.HasZoomControl = True
'
''      Set oReportFormulaFieldDefinations = oReport.FormulaFields
''      Set oReportFF = oReportFormulaFieldDefinations.Item(1)
''      oReportFF.text = "'Factory Information'"
'
'objReport.DiscardSavedData
'objReport.Preview "Doctor_Serial Infromation of '" & cmbDName.text & "'", , , , , 16777216 Or 524288 Or 65536
'
'End Sub
'
'Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyAscii = 13 Then
''       SendKeys Chr(9)
''    End If
'End Sub
'
'Private Sub txtPatientName_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'       SendKeys Chr(9)
'txtPatientName.text = StrConv(txtPatientName.text, vbProperCase)
'    End If
'End Sub
'
'Private Sub txtPhoneNo_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyAscii = 13 Then
''       SendKeys Chr(9)
''    End If
'End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub
