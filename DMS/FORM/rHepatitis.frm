VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rHepatitis 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Lab Report Format [HEPATITIS]"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   DrawWidth       =   2
   Icon            =   "rHepatitis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete_TempTable1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "D E L E T E"
      Height          =   2625
      Left            =   9570
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2820
      Width           =   315
   End
   Begin VB.TextBox txtPat_ID1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2820
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   840
      Width           =   1260
   End
   Begin VB.TextBox txtSN 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2730
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   7365
      Width           =   6870
   End
   Begin VB.ComboBox ComTest_Title 
      DataSource      =   "Adodc7"
      Height          =   315
      Left            =   2760
      TabIndex        =   10
      Top             =   2430
      Width           =   3255
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
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
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8190
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   8550
      Top             =   210
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "8-show S_NAME"
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   8550
      Top             =   225
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "7-show M_CODE"
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8550
      Top             =   225
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "6-S_code lost_focus"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   8550
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "5-flush from Report_All"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "rHepatitis.frx":000C
      Height          =   975
      Left            =   3990
      TabIndex        =   29
      Top             =   1290
      Visible         =   0   'False
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   1720
      _Version        =   393216
      BorderStyle     =   0
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
   Begin VB.CommandButton cmdShow 
      Caption         =   "S&how"
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
      Left            =   8565
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   810
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8550
      Top             =   210
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "4-show S_NAME"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8550
      Top             =   150
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "3-show all pat_info from report_all"
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
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8190
      Width           =   1050
   End
   Begin VB.TextBox txtNote 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "rHepatitis.frx":0021
      Top             =   6540
      Width           =   6825
   End
   Begin VB.TextBox txtS_Name 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4305
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1232
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox txtM_Code 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4290
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtS_Code 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2805
      TabIndex        =   3
      Top             =   1232
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtPat_ID 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2820
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtN_Exam 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2805
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "As Under"
      Top             =   1830
      Width           =   6750
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
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
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
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
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8190
      Width           =   1050
   End
   Begin VB.TextBox txtSpecimen 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2805
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Blood"
      Top             =   1530
      Width           =   3390
   End
   Begin VB.TextBox txtTest_Result 
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   2790
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2820
      Width           =   3315
   End
   Begin VB.TextBox txtUnit 
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   6210
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   2820
      Width           =   3315
   End
   Begin VB.TextBox txtUsed_tech 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2790
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2145
      Width           =   6765
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C&lear"
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
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8190
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8550
      Top             =   210
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
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
      Caption         =   "1"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8550
      Top             =   210
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "2-Show Advance"
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
   Begin MSComCtl2.DTPicker Dt 
      Height          =   285
      Left            =   7320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   825
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      Format          =   66781185
      CurrentDate     =   37114
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1035
      Left            =   2760
      TabIndex        =   31
      Top             =   5490
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   1826
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Note"
      Height          =   195
      Left            =   1140
      TabIndex        =   30
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HEPATITIS"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   435
      TabIndex        =   28
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      Height          =   165
      Left            =   1140
      TabIndex        =   27
      Top             =   6525
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   3780
      TabIndex        =   26
      Top             =   1245
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Code"
      Height          =   195
      Left            =   1095
      TabIndex        =   25
      Top             =   1230
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature of Examination"
      Height          =   195
      Left            =   1080
      TabIndex        =   24
      Top             =   1830
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      Height          =   195
      Left            =   1095
      TabIndex        =   23
      Top             =   855
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   6870
      TabIndex        =   22
      Top             =   855
      Width           =   345
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Specimen"
      Height          =   195
      Left            =   1080
      TabIndex        =   21
      Top             =   1515
      Width           =   750
   End
End
Attribute VB_Name = "rHepatitis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp_Table1 As New ADODB.Recordset
Dim Temp_Table_Helper1 As New ADODB.Recordset
Dim Temp_Table2 As New ADODB.Recordset
Dim Temp_Table_Helper2 As New ADODB.Recordset


Dim Strpat_id1 As String
Dim StrRow_Count As String
Dim StrPat_Type As String
Dim IntPat_ID As Double

Private Sub cmdClear_Click()
     
  '  txtM_Code.Text = ""
'    txtS_Code.Text = ""
'    txtS_Name.Text = ""
    txtSpecimen.text = "Blood"
    txtN_Exam.text = "As Under"
    txtUsed_tech.text = ""
    'txtOpenion = ""
    'txtSample_Rate = ""
    'txtCut_of_Rate = ""
'    txtAnti_Virus1 = ""
'    txtAnti_Virus2 = ""
'    txtAnti_Virus3 = ""
    Dt.value = Date
    'txtNote = ""
 '   ComTest_Title.Text = ""
    txtTest_Result.text = ""
    txtUnit.text = ""
    'txtSN.Text = ""
    txtPat_ID.text = ""
    txtPat_ID1.text = ""
    txtNote.text = ""
    Temp_rst1
    ComTest_Title.Clear
    DataGrid1.Visible = False
    txtPat_ID1.SetFocus
    
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    If txtPat_ID = "" Then Exit Sub
'    If txtS_Code = "" Then Exit Sub
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
    'Del_All_Report_All_TempRst1
    Del_Report
    
    txtTest_Result = ""
    txtUnit.text = ""
    txtNote = ""
    Temp_rst1
    'txtNote.Text = ""
    txtUsed_tech = ""
    txtNote = ""
    txtPat_ID = ""
    txtPat_ID1 = ""
    ComTest_Title.Clear
    txtPat_ID1.SetFocus
    End If
End Sub
Private Sub Del_All_Report_All_TempRst1()
   
'    Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
 '   While Temp_Table1.EOF = False
            cmd.CommandText = "exec Report_All_Delete2 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
            cmd.Execute
   '         Temp_Table1.MoveNext
  '  Wend
    con.Close
End Sub

Private Sub cmdDelete_TempTable1_Click()
If ComTest_Title = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
    If Temp_Table1.RecordCount <= 0 Then Exit Sub
    
    If Trim(ComTest_Title.text) = "" Then
        MsgBox "You didn't select the Test Name"
        DataGrid2.SetFocus
        Exit Sub
    Else
        Dim Strmsg As String
        Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
        DelReport_All_TempRst1
        Temp_Table1.Delete
        ComTest_Title = ""
        txtTest_Result = ""
        txtUnit = ""
        txtRef_Range = ""
        End If
        
    End If
End Sub

Private Sub CmdPreview_Click()

    CRViewer1_MODE = 19
    Viewer.Show vbModal
    
End Sub

Private Sub cmdPrint_Click()
    '==========direct print==========================
            
            Dim Report19 As New Hepatities
            Dim StrPat_ID As String
           
            Dim strM_Code As String
            Dim strS_Code As String
            
            StrPat_ID = rHepatitis.txtPat_ID
            StrPat_ID_R = StrPat_ID
            
            strM_Code = rHepatitis.txtM_Code
            strS_Code = rHepatitis.txtS_Code
     
            '--------------------------------------------------------------------
            Report19.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report19.FormulaFields.Item(2).text = Chr(34) & "Patient ID" & Chr(34)
            Report19.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report19.FormulaFields.Item(4).text = Chr(34) & "Delivered Date" & Chr(34)
            Report19.FormulaFields.Item(5).text = Chr(34) & "Patient Name" & Chr(34)
            Report19.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report19.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report19.FormulaFields.Item(8).text = Chr(34) & "Refd. by" & Chr(34)
            '--------------------------------------------------------------------
            Report19.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report19.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report19.FormulaFields.Item(11).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report19.Text1.SetText StDoc_Name
            
            Report19.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report19.Database.SetDataSource rs
           
            Report19.PrintOut (False)
            rs.Close
            Call cmdClear_Click
            txtPat_ID1.SetFocus
    '====================================
End Sub

Private Sub cmdSave_Click()
'-----validation check---------------------
    If Trim(txtPat_ID) = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID.SetFocus
        Exit Sub
    End If
'    If Len(Trim(txtS_Code)) = 0 Then
'        MsgBox "Test Code mandatory"
'        txtS_Code.SetFocus
'        Exit Sub
'    End If
    
'-----end validation check--------------------------------------------------
    Adodc2.connectionstring = strcn.Connection
    'Adodc2.RecordSource = "select * from Report_All where pat_id='" & Trim(txtPat_ID.Text) & "'"
    Adodc2.RecordSource = "Report_All_SELECT2 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        Del_All_Report_All_TempRst1
        InsReport_All_TempRst1
        MsgBox "Updated Successfully"
'       UpdReport_All
    Else
'       InsReport_All
        InsReport_All_TempRst1
       MsgBox "Inserted Successfully"
    End If
    
    Temp_rst1
    
    cmdPrint.SetFocus
End Sub
Private Sub cmdShow_Click()

    If txtPat_ID1.text = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID = ""
        txtPat_ID1.SetFocus
        Exit Sub
    End If
'===for show data in Datagrid1=============
    Adodc1.connectionstring = strcn.Connection
'    Adodc1.RecordSource = " select a.m_code,a.s_code,(select s_name from test_info_sub b where a.s_code=b.s_code) as s_name from pat_info_sub1 a where pat_id='" + txtPat_ID + "'"
    Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.Visible = True
'        DataGrid1.Columns(2).Width = 50250
'        DataGrid1.Columns(0).Caption = "Group Code"
'        DataGrid1.Columns(1).Caption = "Test Code"
'        DataGrid1.Columns(2).Caption = "   Name of Test"
    Else
            DataGrid1.Visible = False
            MsgBox "Invalid Patient ID"
            txtPat_ID = ""
            txtPat_ID1 = ""
            txtPat_ID1.SetFocus
            Exit Sub
    End If
'===============================================
End Sub
Private Sub ComTest_Title_LostFocus()
    If ComTest_Title = "" Then
        cmdSave.SetFocus
        Exit Sub
    End If
    'GetResult
    GetS_Code
End Sub
Private Sub DataGrid1_DblClick()
'    txtM_Code.Text = DataGrid1.Columns(0)
    txtS_Code.text = DataGrid1.Columns(1)
    StrSub_Code = Me.txtS_Code
    txtS_Name.text = DataGrid1.Columns(2)
    txtSpecimen.SetFocus
    DataGrid1.Visible = False
    
End Sub

Private Sub DataGrid2_DblClick()
On Error Resume Next
ComTest_Title = DataGrid2.Columns(0)
txtTest_Result.text = DataGrid2.Columns(1)
txtUnit.text = DataGrid2.Columns(2)
    
End Sub
Private Sub Form_Click()
    DataGrid1.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    If rHepatitis.DataGrid1.Visible = True Then
        rHepatitis.DataGrid1.Visible = False
    Else
        Unload Me
    End If
    End If
End Sub

Private Sub Form_Load()



    Adodc7.connectionstring = strcn.Connection
    Adodc7.RecordSource = "exec m_name_select 1,'" + "HEPATITIS PROFILE" + "'"
    Adodc7.Refresh
    If Adodc7.Recordset.RecordCount > 0 Then
        txtM_Code = Adodc7.Recordset!m_code
    Else
        MsgBox "Inserted incurrect head name, first you have to insert currect name from TEST INFORMATION form then open this screen again"
        txtPat_ID.Enabled = False
        cmdSave.Enabled = False
    End If

    'GetTestTitle
    
    Dt.value = Date
    
    Temp_rst1
    
    StrScreenName = "Hepatitis"
    Flush_Font_Type
    
End Sub

Private Sub txtCut_of_Rate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtN_Exam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtOpenion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtPat_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub
Private Sub txtPat_ID_LostFocus()
'Pat_Paid

If Len(Trim(txtPat_ID.text)) = 0 Then Exit Sub
    
Adodc5.connectionstring = strcn.Connection
Adodc5.RecordSource = "exec Pro_FLUSH 6," & txtPat_ID & ""
Adodc5.Refresh
If Adodc5.Recordset.RecordCount > 0 Then

Else
    '===for show data in Datagrid1=============
                Adodc1.connectionstring = strcn.Connection
                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
                Adodc1.Refresh
               If Adodc1.Recordset.RecordCount > 0 Then
                    DataGrid1.Visible = True
'                    DataGrid1.Columns(2).Width = 1000
'                    DataGrid1.Columns(0).Caption = "Group Code"
'                    DataGrid1.Columns(1).Caption = "Test Code"
'                    DataGrid1.Columns(2).Caption = "   Name of Test"
                Else
                    DataGrid1.Visible = False
                    MsgBox "Invalied ID"
                    txtPat_ID = ""
                    txtPat_ID.SetFocus
                    
                End If
        '===============================================
End If
End Sub
Private Sub InsReport_All_TempRst1()
    
    Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    While Temp_Table1.EOF = False
    
          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
            "','" + txtM_Code + _
            "','" + txtS_Code + _
            "','" + ChkForQuote(txtSpecimen) + _
            "','" + ChkForQuote(txtN_Exam) + _
            "','" + ChkForQuote(txtUsed_tech) + _
            "','" + ChkForQuote(Temp_Table1!Test_Name) + _
            "','" + ChkForQuote(Temp_Table1!Test_result) + _
            "','" + ChkForQuote(Temp_Table1!unit) + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + ChkForQuote(txtSN.text) + _
            "','" + ChkForQuote(txtNote.text) + _
            "','" + u_id + _
            "','" + Format(Dt, "yyyy-mm-dd") + _
            "','" + "" + _
            "','" + txtPat_ID1 + "'"
            cmd.Execute
            Temp_Table1.MoveNext
    Wend
    con.Close
End Sub

Private Sub txtPat_ID1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtPat_ID1_LostFocus()


If txtPat_ID1 = "" Then Exit Sub
       
    Search_Patient_Type
    
    If StrRow_Count > "1" Then
        
            Dim Patmsg As String
            Patmsg = MsgBox("Do you want to update Inside Patient's information ? ", vbQuestion + vbYesNo)
            If Patmsg = vbYes Then
                StrPat_Type = "0"
                Srch_Pat_ID
            Else
                StrPat_Type = "1"
                Srch_Pat_ID
                
            End If
    Else
            Srch_Pat_ID1
    End If
    
   
    txtPat_ID = IntPat_ID
    
   
    If IntPat_ID = 0 Then
        MsgBox "Invalid ID, Try again"
        txtPat_ID = ""
        txtPat_ID1 = ""
        txtPat_ID1.SetFocus
        Exit Sub
    End If

'----------------------------

txtS_Name.ForeColor = vbBlack 'for show s_name
'txtS_Name = ""                'for show s_name

'If Len(txtS_Code.Text) = 0 Then Exit Sub
If cmdSave.Enabled = False Then Exit Sub

StrSub_Code = txtS_Code.text

Temp_rst1

'If Trim(txtPat_ID1) = "" Then
'    MsgBox "Patient ID mandatory"
'    txtPat_ID1.SetFocus
'    Exit Sub
'End If

'If Len(Trim(txtS_Code)) = 0 Then Exit Sub
    

    
    'for flush patient information
     Adodc6.connectionstring = strcn.Connection
     'Adodc3.RecordSource = "exec Pat_Info_SELECT 4,'" + txtPat_ID + "'"
     Adodc6.RecordSource = "exec Report_All_SELECT3 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
     Adodc6.Refresh
        
    If Adodc6.Recordset.RecordCount > 0 Then
         txtSpecimen = Adodc6.Recordset!Field1
         txtN_Exam = Adodc6.Recordset!Field2
         txtUsed_tech = Adodc6.Recordset!Field3
        ' Me.ComTest_Title = Adodc6.Recordset!Field4
         'Me.txtTest_Result = Adodc6.Recordset!Field5
         'txtUnit = Adodc6.Recordset!Field6
         txtSN.text = Adodc6.Recordset!Field14
         txtNote = Adodc6.Recordset!Field15
         Dt.value = Adodc6.Recordset!Dt
         
         
         
         '/////////show in Temp_rst1//////////////
        con.connectionstring = strcn.Connection
        con.Open
        Temp_Table_Helper1.Open "select * from report_all where pat_id='" + txtPat_ID + "' and s_code='" + txtS_Code + "'and m_code='" + txtM_Code + "'", con
        
          While Temp_Table_Helper1.EOF = False
                Temp_Table1.AddNew
                Temp_Table1!Test_Name = Temp_Table_Helper1!Field4
                Temp_Table1!Test_result = Temp_Table_Helper1!Field5
                Temp_Table1!unit = Temp_Table_Helper1!Field6
                'Temp_Table1!ref_range = Temp_Table_Helper1!Field7
                Temp_Table_Helper1.MoveNext
          Wend
        DataGrid2.Refresh
        Temp_Table_Helper1.Close
        con.Close
'/////////end show in Temp_rst1////////////////////////////

         

    Else
        '===for show data in Datagrid1=============
                Adodc1.connectionstring = strcn.Connection
                'Adodc1.RecordSource = "select a.m_code,a.s_code,(select s_name from test_info_sub b where a.s_code=b.s_code) as s_name from pat_info_sub1 a where pat_id='" + txtPat_ID + "'"
                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
                Adodc1.Refresh
        
                If Adodc1.Recordset.RecordCount > 0 Then
                    DataGrid1.Visible = True
                    DataGrid1.Columns(2).Width = 5270
                    DataGrid1.Columns(0).Caption = "Group Code"
                    DataGrid1.Columns(1).Caption = "Test Code"
                    DataGrid1.Columns(2).Caption = "   Name of Test"
                Else
                    DataGrid1.Visible = False
                End If
        '===============================================
       End If

'GetTestTitle


Call BindTestName

End Sub

'Private Sub Del_All_Report_All_TempRst1()
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'            cmd.CommandText = "exec Report_All_Delete2 1,'" + Trim(txtPat_ID.Text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
'            cmd.Execute
'    con.Close
'End Sub
Private Sub txtS_Code_Change()
    '--------FOR S_NAME-------------------------
        Adodc8.connectionstring = strcn.Connection
        Adodc8.RecordSource = "exec S_name_select 1,'" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
        Adodc8.Refresh
        If Adodc8.Recordset.RecordCount > 0 Then
            txtS_Name.ForeColor = vbBlack
            txtS_Name = Adodc8.Recordset!s_name
        Else
            txtS_Name.ForeColor = vbRed
            txtS_Name = " Invalied Test code"
        End If
'--------END---------------------------------
End Sub

Private Sub txtS_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtS_Code_LostFocus()
txtS_Name.ForeColor = vbBlack 'for show s_name
'txtS_Name = ""                'for show s_name

'If Len(txtS_Code.Text) = 0 Then Exit Sub
If cmdSave.Enabled = False Then Exit Sub

StrSub_Code = txtS_Code.text

Temp_rst1

'If Trim(txtPat_ID1) = "" Then
'    MsgBox "Patient ID mandatory"
'    txtPat_ID1.SetFocus
'    Exit Sub
'End If

'If Len(Trim(txtS_Code)) = 0 Then Exit Sub
    

    
    'for flush patient information
     Adodc6.connectionstring = strcn.Connection
     'Adodc3.RecordSource = "exec Pat_Info_SELECT 4,'" + txtPat_ID + "'"
     Adodc6.RecordSource = "exec Report_All_SELECT3 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
     Adodc6.Refresh
        
    If Adodc6.Recordset.RecordCount > 0 Then
         txtSpecimen = Adodc6.Recordset!Field1
         txtN_Exam = Adodc6.Recordset!Field2
         txtUsed_tech = Adodc6.Recordset!Field3
        ' Me.ComTest_Title = Adodc6.Recordset!Field4
         'Me.txtTest_Result = Adodc6.Recordset!Field5
         'txtUnit = Adodc6.Recordset!Field6
         txtSN.text = Adodc6.Recordset!Field14
         txtNote = Adodc6.Recordset!Field15
         Dt.value = Adodc6.Recordset!Dt
         
         
         
         '/////////show in Temp_rst1//////////////
        con.connectionstring = strcn.Connection
        con.Open
        Temp_Table_Helper1.Open "select * from report_all where pat_id='" + txtPat_ID + "' and s_code='" + txtS_Code + "'and m_code='" + txtM_Code + "'", con
        
          While Temp_Table_Helper1.EOF = False
                Temp_Table1.AddNew
                Temp_Table1!Test_Name = Temp_Table_Helper1!Field4
                Temp_Table1!Test_result = Temp_Table_Helper1!Field5
                Temp_Table1!unit = Temp_Table_Helper1!Field6
                'Temp_Table1!ref_range = Temp_Table_Helper1!Field7
                Temp_Table_Helper1.MoveNext
          Wend
        DataGrid2.Refresh
        Temp_Table_Helper1.Close
        con.Close
'/////////end show in Temp_rst1////////////////////////////

         

    Else
        '===for show data in Datagrid1=============
                Adodc1.connectionstring = strcn.Connection
                'Adodc1.RecordSource = "select a.m_code,a.s_code,(select s_name from test_info_sub b where a.s_code=b.s_code) as s_name from pat_info_sub1 a where pat_id='" + txtPat_ID + "'"
                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
                Adodc1.Refresh
        
                If Adodc1.Recordset.RecordCount > 0 Then
                    DataGrid1.Visible = True
                    DataGrid1.Columns(2).Width = 5270
                    DataGrid1.Columns(0).Caption = "Group Code"
                    DataGrid1.Columns(1).Caption = "Test Code"
                    DataGrid1.Columns(2).Caption = "   Name of Test"
                Else
                    DataGrid1.Visible = False
                End If
        '===============================================
       End If

'GetTestTitle

End Sub

Private Sub txtSample_Rate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub txtSpecimen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub
Private Sub txtUnit_LostFocus()

'----------------check--------
Dim Check As Integer
Check = 0
If Temp_Table1.RecordCount > 0 Then
    Temp_Table1.MoveFirst

        While Temp_Table1.EOF = False

            If Temp_Table1!Test_Name = ComTest_Title Then
                Check = 1
            End If
    Temp_Table1.MoveNext
        Wend
    If Check = 1 Then
        MsgBox "This Test Name already exists"
        Check = 0
        ComTest_Title.SetFocus
        Exit Sub
    End If
'    Temp_Table.MoveFirst
End If


'--------------end check-----

'+++to insert into TEMPORARY RECORDSET "Temp_rst1"++++
        Temp_Table1.AddNew
        Temp_Table1!Test_Name = ComTest_Title
        Temp_Table1!Test_result = txtTest_Result
        Temp_Table1!unit = txtUnit
        'Temp_Table1!ref_range = txtRef_Range
        DataGrid2.Refresh
'+++++++++++++++++++++++++++++++++++++++

txtTest_Result = ""
ComTest_Title = ""
txtUnit = ""
ComTest_Title.SetFocus

End Sub

Private Sub txtUsed_tech_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub
Private Sub GetTestTitle()
StrSub_Code = Me.txtS_Code
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select1 1,'" + "04" + "'"
  Adodc7.RecordSource = "exec Flush_Test_Result 1,'" & txtM_Code.text & "','" & txtM_Code & "','" & StrSub_Code & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
       Do Until Adodc7.Recordset.EOF
          ComTest_Title.AddItem Adodc7.Recordset!Test_Name
       Adodc7.Recordset.MoveNext
       Loop
    End If
End Sub
Private Sub GetResult()
  Adodc7.connectionstring = strcn.Connection
  'Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Title.Text + "','" + "04" + "'"
  Adodc7.RecordSource = "exec Flush_Test_Result 1,'" & txtM_Code.text & "','" & txtM_Code & "','" & StrSub_Code & "'"
  Adodc7.Refresh
    If Adodc7.Recordset.RecordCount > 0 Then
        txtTest_Result.text = Adodc7.Recordset!Test_result
        txtUnit.text = Adodc7.Recordset!unit
        txtUsed_tech = Adodc7.Recordset!ref_range
        txtNote.text = Adodc7.Recordset!others
    End If
End Sub
Private Sub Pat_Paid()
    Adodc5.connectionstring = strcn.Connection
    Adodc5.RecordSource = "exec Select_Paid 1,'" + txtPat_ID + "'"
    Adodc5.Refresh
    If Adodc5.Recordset.RecordCount > 0 Then
    Else
        txtPat_ID = ""
        MsgBox " Patient could not paid"
    End If
    
End Sub


Private Sub Search_Patient_Type()

    StrRow_Count = "1"
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Pat_Type 1,'" & txtPat_ID1.text & "'", con
    If My_Rst.EOF = False Then
    
        StrRow_Count = My_Rst!Row_Count
        'MsgBox StrRow_Count
    End If
    
    con.Close
    
End Sub
Private Sub Srch_Pat_ID()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Pat_ID 1,'" & txtPat_ID1.text & "','" & StrPat_Type & "'", con
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!pat_id2
  '      MsgBox IntPat_ID
    End If
    con.Close
    
End Sub
Private Sub Srch_Pat_ID1()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Pat_ID1 1,'" & txtPat_ID1.text & "'", con
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!pat_id2
 '       MsgBox IntPat_ID
    End If
    con.Close
    
End Sub

Public Sub Temp_rst1()
    
    Set Temp_Table1 = New ADODB.Recordset
    With Temp_Table1
        .Fields.Append "Test_Name", adVarChar, 500
        .Fields.Append "Test_Result", adVarChar, 500
        .Fields.Append "Unit", adVarChar, 500
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set DataGrid2.DataSource = Temp_Table1
    
    DataGrid2.Columns(0).DataField = "Test_Name"
    DataGrid2.Columns(1).DataField = "Test_Result"
    DataGrid2.Columns(2).DataField = "Unit"
    DataGrid2.ReBind
    DataGrid2.Refresh
    
    DataGrid2.Columns(0).Width = 3270.047
    DataGrid2.Columns(1).Width = 3284.788
    DataGrid2.Columns(2).Width = 1950.236

End Sub
Private Sub DelReport_All_TempRst1()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "" + "'"
            cmd.Execute
    con.Close
End Sub
Private Sub BindTestName()
    On Error GoTo err_loop
       ComTest_Title.Clear
       con.connectionstring = strcn.Connection
       con.Open
       rs.Open "exec GetTestName '" & Trim(txtM_Code.text) & "','" & Trim(txtPat_ID.text) & "'", con

       If rs.EOF = False Then
          Do Until rs.EOF
            ComTest_Title.AddItem rs!Test_Name
          rs.MoveNext
          Loop
       End If
       rs.Close
       con.Close
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub


Private Sub GetS_Code()

  Adodc7.connectionstring = strcn.Connection
  'Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Title + "','" + "01" + "'"
  Adodc7.RecordSource = "exec test_Result_Select8 '" & txtPat_ID & "','" & ComTest_Title.text & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    
    'ComTest_Name.AddItem Adodc7.Recordset!test_result
    txtTest_Result = Adodc7.Recordset!Test_result
    txtUnit = Adodc7.Recordset!unit
    txtUsed_tech = Adodc7.Recordset!ref_range
    txtNote = Adodc7.Recordset!others
    'txtNormal_Value = Adodc7.Recordset!ref_range
    
    Adodc7.Recordset.MoveNext
    Loop
    'txtUnit = Adodc7.Recordset!unit
    'txtUnit = Adodc7.Recordset!unit
    'txtNormal_Value = Adodc7.Recordset!ref_range   'ref_range
    End If
End Sub
Private Sub Del_Report()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Del_Report 1,'" + Trim(txtPat_ID.text) + "'"
            cmd.Execute
    con.Close
End Sub

