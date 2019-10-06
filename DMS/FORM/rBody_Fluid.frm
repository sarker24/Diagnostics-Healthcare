VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rBody_Fluid 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Lab Report Format [BODY FLUID]"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   DrawWidth       =   2
   Icon            =   "rBody_Fluid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPat_ID1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2250
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   510
      Width           =   1260
   End
   Begin VB.TextBox txtSN 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   450
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   7575
      Width           =   8640
   End
   Begin VB.TextBox txtUnit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4770
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   2910
      Width           =   4320
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "rBody_Fluid.frx":000C
      Height          =   840
      Left            =   3120
      TabIndex        =   21
      Top             =   1290
      Visible         =   0   'False
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   1482
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
      Left            =   8910
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   690
      Width           =   1170
   End
   Begin VB.TextBox txtS_Name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3705
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   5115
   End
   Begin VB.TextBox txtN_Exam 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2235
      TabIndex        =   8
      Text            =   "As Under"
      Top             =   1575
      Width           =   6600
   End
   Begin VB.TextBox txtPh_Exam 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   450
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "PHYSICAL EXAMINATION"
      Top             =   2160
      Width           =   3405
   End
   Begin VB.TextBox txtNote 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   450
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   6750
      Width           =   8640
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   2880
      TabIndex        =   15
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   6030
      TabIndex        =   18
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   8130
      TabIndex        =   20
      Top             =   8190
      Width           =   1050
   End
   Begin VB.TextBox txtTest_Result 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   450
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2910
      Width           =   4290
   End
   Begin VB.CommandButton cmdClear 
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
      Left            =   7080
      TabIndex        =   19
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdPreview 
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
      Left            =   4980
      TabIndex        =   17
      Top             =   8190
      Width           =   1050
   End
   Begin VB.ComboBox ComTest_Name 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc7"
      Height          =   315
      Left            =   420
      TabIndex        =   10
      Top             =   2460
      Width           =   3405
   End
   Begin VB.TextBox txtS_Code 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2235
      TabIndex        =   3
      Top             =   1020
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtM_Code 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3705
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   345
   End
   Begin VB.TextBox txtPat_ID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2235
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   735
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtSpecimen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2235
      TabIndex        =   7
      Top             =   1305
      Width           =   3390
   End
   Begin VB.CommandButton cmdPrint 
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
      Left            =   3930
      TabIndex        =   16
      Top             =   8190
      Width           =   1050
   End
   Begin MSComCtl2.DTPicker Dt 
      Height          =   285
      Left            =   7575
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   690
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      Format          =   65273857
      CurrentDate     =   37114
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   9240
      Top             =   90
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
      Caption         =   "Adodc7"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   9240
      Top             =   90
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
      Caption         =   "Adodc4"
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
      Left            =   9240
      Top             =   90
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
      Caption         =   "Adodc5"
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
      Left            =   9240
      Top             =   90
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   9240
      Top             =   90
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
      Caption         =   "Adodc8"
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   9240
      Top             =   90
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
      Caption         =   "9-S_NAME"
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   9240
      Top             =   90
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
      Caption         =   "Adodc11"
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   9240
      Top             =   90
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
      Caption         =   "Adodc10"
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
      Left            =   9240
      Top             =   90
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
      Caption         =   "3-pat_ID"
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
      Left            =   9240
      Top             =   90
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
      Left            =   9240
      Top             =   90
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Note"
      Height          =   195
      Left            =   450
      TabIndex        =   29
      Top             =   7320
      Width           =   915
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Specimen"
      Height          =   195
      Left            =   420
      TabIndex        =   28
      Top             =   1350
      Width           =   750
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BODY FLUID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4530
      TabIndex        =   27
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   3210
      TabIndex        =   26
      Top             =   1050
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Code"
      Height          =   195
      Left            =   420
      TabIndex        =   25
      Top             =   1065
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature of Examination"
      Height          =   195
      Left            =   420
      TabIndex        =   24
      Top             =   1635
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      Height          =   195
      Left            =   420
      TabIndex        =   23
      Top             =   780
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   7020
      TabIndex        =   22
      Top             =   750
      Width           =   345
   End
End
Attribute VB_Name = "rBody_Fluid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Temp_Table1 As New ADODB.Recordset
'Dim Temp_Table_Helper1 As New ADODB.Recordset
'Dim Temp_Table2 As New ADODB.Recordset
'Dim Temp_Table_Helper2 As New ADODB.Recordset
'Dim Temp_Table3 As New ADODB.Recordset
'Dim Temp_Table_Helper3 As New ADODB.Recordset

Dim Strpat_id1 As String
Dim StrRow_Count As String
Dim StrPat_Type As String
Dim IntPat_ID As Double


Private Sub cmdClear_Click()
'    Temp_rst1
'    Temp_rst2
'    Temp_rst3
    txtSpecimen = ""
    txtTest_Result = ""
    txtUnit = ""
    txtNote.text = ""
    txtSN.text = ""
    
    If DataGrid1.Visible = True Then
        DataGrid1.Visible = False
    End If
    
    txtPat_ID1 = ""
    txtPat_ID = ""
    txtS_Code = ""
    ComTest_Name.Clear
    txtPat_ID1.SetFocus
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    If cmdSave.Enabled = False Then Exit Sub

    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
    'Del_All_Report_All_TempRst1
    Del_Report
    Clearscreen
    End If
End Sub
Private Sub cmdDelete_TempTable1_Click()

    If ComTest_Name = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
    If Temp_Table1.RecordCount <= 0 Then Exit Sub
    
    If Trim(ComTest_Name.text) = "" Then
        MsgBox "You didn't select the the Test Name"
        DataGrid2.SetFocus
        Exit Sub
    Else
        Dim Strmsg As String
        Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
        'DelReport_All_TempRst1
        'Temp_Table1.Delete
        ComTest_Name = ""
        txtTest_Result = ""
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_TempTable2_Click()
    If ComTest_Name1 = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
    If Temp_Table2.RecordCount <= 0 Then Exit Sub
    
    If Trim(ComTest_Name.text) = "" Then
        MsgBox "You didn't select the the Test Name"
        DataGrid3.SetFocus
        Exit Sub
    Else
        Dim Strmsg As String
        Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
'        DelReport_All_TempRst2
'        Temp_Table2.Delete
        ComTest_Name1 = ""
        txtTest_Result1 = ""
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_TempTable3_Click()
    If ComTest_Name2 = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
    If Temp_Table3.RecordCount <= 0 Then Exit Sub
    
    If Trim(ComTest_Name2.text) = "" Then
        MsgBox "You didn't select the the Test Name"
        DataGrid4.SetFocus
        Exit Sub
    Else
        Dim Strmsg As String
        Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
'        DelReport_All_TempRst3
'        Temp_Table3.Delete
        ComTest_Name2 = ""
        txtTest_Result2 = ""
        
        End If
        
    End If
End Sub

Private Sub CmdPreview_Click()
    CRViewer1_MODE = 23
    Viewer.Show vbModal
End Sub

Private Sub cmdPrint_Click()
    '==========direct print==========================
            
            Dim Report23 As New BodyFluid
            Dim StrPat_ID As String
            
            
            
            Dim strM_Code As String
            Dim strS_Code As String
            
            StrPat_ID = rBody_Fluid.txtPat_ID
            StrPat_ID_R = StrPat_ID
            
            strM_Code = rBody_Fluid.txtM_Code
            strS_Code = rBody_Fluid.txtS_Code
            
            '--------------------------------------------------------------------
            Report23.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report23.FormulaFields.Item(2).text = Chr(34) & "Patient ID" & Chr(34)
            Report23.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report23.FormulaFields.Item(4).text = Chr(34) & "Delivered Date" & Chr(34)
            Report23.FormulaFields.Item(5).text = Chr(34) & "Patient Name" & Chr(34)
            Report23.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report23.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report23.FormulaFields.Item(8).text = Chr(34) & "Refd. by" & Chr(34)
            '--------------------------------------------------------------------
            Report23.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report23.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report23.FormulaFields.Item(11).text = Chr(34) & "Checked By" & Chr(34)
            
            Call Flush_Doc_Name
            Report23.Text1.SetText StDoc_Name
            
            Report23.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report23.Database.SetDataSource rs
            
            Report23.PrintOut (False)
            rs.Close
            Call cmdClear_Click
            txtPat_ID1.SetFocus
    '====================================
End Sub

Private Sub cmdSave_Click()
'-----validation check---------------------
    If Trim(txtPat_ID) = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID1.SetFocus
        Exit Sub
    End If
    
'    If Len(Trim(txtS_Code)) = 0 Then
'        MsgBox "Test Code mandatory"
'        txtS_Code.SetFocus
'        Exit Sub
'    End If
'-----end validation check--------------------------------------------------

''\\\\\\\\\\INSERT and UPDATE from Temp_rst1\\\\\\\\\\\\\
    Adodc2.connectionstring = strcn.Connection
'    Adodc2.RecordSource = "Report_All_SELECT3 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc2.RecordSource = "Report_All_SELECT 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "3" + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        Del_All_Report_All_TempRst1
        InsReport_All_TempRst1

        MsgBox "Updated"
    Else
        InsReport_All_TempRst1

        MsgBox "Inserted"
    End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
''\\\\\\\\\\INSERT and UPDATE from Temp_rst2\\\\\\\\\\\\\

DataGrid1.Visible = False
cmdPrint.SetFocus

End Sub
Private Sub cmdShow_Click()
If cmdSave.Enabled = False Then Exit Sub

        If txtPat_ID1.text = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID1.SetFocus
        Exit Sub
    End If
'===for show data in Datagrid1=============
    Adodc1.connectionstring = strcn.Connection

    Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.Visible = True
        DataGrid1.Columns(2).Width = 5250
        DataGrid1.Columns(0).Caption = "Group Code"
        DataGrid1.Columns(1).Caption = "Test Code"
        DataGrid1.Columns(2).Caption = "   Name of Test"
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
Private Sub ComResult1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub ComResult2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub ComTest_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub DataGrid2_DblClick()
    ComTest_Name.text = DataGrid2.Columns(0)
    txtTest_Result.text = DataGrid2.Columns(1)
    
End Sub
Private Sub DataGrid3_DblClick()
    ComTest_Name1.text = DataGrid3.Columns(0)
    txtTest_Result1.text = DataGrid3.Columns(1)
End Sub
Private Sub DataGrid4_DblClick()
    ComTest_Name2.text = DataGrid4.Columns(0)
    txtTest_Result2.text = DataGrid4.Columns(1)
End Sub

Private Sub ComTest_Name_LostFocus()
If ComTest_Name = "" Then
    cmdSave.SetFocus
End If
    GetS_Code
    GetSpecimen
   ' GetResult
End Sub

Private Sub Form_Click()
    DataGrid1.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    If rBody_Fluid.DataGrid1.Visible = True Then
        rBody_Fluid.DataGrid1.Visible = False
    Else
        Unload Me
    End If
    End If
End Sub

Private Sub Form_Load()
    
    Adodc8.connectionstring = strcn.Connection
    Adodc8.RecordSource = "exec m_name_select 2,'" + "BODY FLUID" + "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
        txtM_Code = Adodc8.Recordset!m_code
    Else
        MsgBox "Inserted incurrect group name, first you have to insert currect name from TEST INFORMATION form then open this screen again"
        txtPat_ID.Enabled = False
        cmdSave.Enabled = False
    End If


    Dt.value = Date
  
StrScreenName = "Body Fluid"
Flush_Font_Type

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Temp_Table1 = Nothing

End Sub
Private Sub txtN_Exam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub txtPat_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub DataGrid1_DblClick()
    'txtM_Code.Text = DataGrid1.Columns(0)
    txtS_Code.text = DataGrid1.Columns(1)
    txtS_Name.text = DataGrid1.Columns(2)
    txtSpecimen.SetFocus
    DataGrid1.Visible = False
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
        'txtDummy_Pat_ID = ""
        txtPat_ID1.SetFocus
        Exit Sub
    End If

'----------------------------

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
                    'DataGrid1.Visible = True
                    DataGrid1.Columns(2).Width = 5270
                    DataGrid1.Columns(0).Caption = "Group Code"
                    DataGrid1.Columns(1).Caption = "Test Code"
                    DataGrid1.Columns(2).Caption = "   Name of Test"
                Else
                    DataGrid1.Visible = False
                    MsgBox "Invalied Patient ID"
                    txtPat_ID = ""
                    txtPat_ID1.SetFocus
                    
                End If
        '===============================================
End If

'-------------------------------------
If cmdSave.Enabled = False Then Exit Sub

If txtPat_ID1 = "" Then
    MsgBox "Patient ID mandatory"
    txtPat_ID1.SetFocus
    Exit Sub
End If

'If Len(Trim(txtS_Code)) = 0 Then Exit Sub
         
    Adodc6.connectionstring = strcn.Connection
    Adodc6.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "3" + "'"
    Adodc6.Refresh
    
    If Adodc6.Recordset.RecordCount > 0 Then
        DataGrid1.Visible = False
    
        txtSpecimen = Adodc6.Recordset!Field1
'         txtSpecimen_dt_Time = Adodc6.Recordset!Field2
        txtN_Exam = Adodc6.Recordset!Field2
'         txtUsed_tech = Adodc6.Recordset!Field3
        Dt.value = Adodc6.Recordset!Dt

'++++++++++for show feild5 to txtNote FROM TYPE 3 +++++++++++
    Adodc8.connectionstring = strcn.Connection
    Adodc8.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "3" + "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
        ComTest_Name.text = Adodc8.Recordset!Field3
        txtTest_Result.text = Adodc8.Recordset!Field4
        txtUnit.text = Adodc8.Recordset!Field5
        txtSN.text = Adodc8.Recordset!Field14
        txtNote = Adodc8.Recordset!Field15
    End If

    Else
    '===for show data in Datagrid1=============
                Adodc1.connectionstring = strcn.Connection
                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
                Adodc1.Refresh

                If Adodc1.Recordset.RecordCount > 0 Then
                    'DataGrid1.Visible = True
                    DataGrid1.Columns(2).Width = 5270
                    DataGrid1.Columns(0).Caption = "Group Code"
                    DataGrid1.Columns(1).Caption = "Test Code"
                    DataGrid1.Columns(2).Caption = "   Name of Test"
                Else
                    DataGrid1.Visible = False
                End If
'===============================================
    End If

Call BindTestName


End Sub

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


'If Len(txtS_Code.Text) = 0 Then Exit Sub
'If cmdSave.Enabled = False Then Exit Sub
'
'If txtPat_ID1 = "" Then
'    MsgBox "Patient ID mandatory"
'    txtPat_ID1.SetFocus
'    Exit Sub
'End If
'
''If Len(Trim(txtS_Code)) = 0 Then Exit Sub
'
'    Adodc6.connectionstring = strcn.Connection
'    Adodc6.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "3" + "'"
'    Adodc6.Refresh
'
'    If Adodc6.Recordset.RecordCount > 0 Then
'        DataGrid1.Visible = False
'
'        txtSpecimen = Adodc6.Recordset!Field1
''         txtSpecimen_dt_Time = Adodc6.Recordset!Field2
'        txtN_Exam = Adodc6.Recordset!Field2
''         txtUsed_tech = Adodc6.Recordset!Field3
'        Dt.value = Adodc6.Recordset!Dt
'
''++++++++++for show feild5 to txtNote FROM TYPE 3 +++++++++++
'    Adodc8.connectionstring = strcn.Connection
'    Adodc8.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "3" + "'"
'    Adodc8.Refresh
'    If Adodc8.Recordset.RecordCount > 0 Then
'        ComTest_Name.Text = Adodc8.Recordset!Field3
'        txtTest_Result.Text = Adodc8.Recordset!Field4
'        txtUnit.Text = Adodc8.Recordset!Field5
'        txtSN.Text = Adodc8.Recordset!Field14
'        txtNote = Adodc8.Recordset!Field15
'    End If
'
'    Else
'    '===for show data in Datagrid1=============
'                Adodc1.connectionstring = strcn.Connection
'                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
'                Adodc1.Refresh
'
'                If Adodc1.Recordset.RecordCount > 0 Then
'                    DataGrid1.Visible = True
'                    DataGrid1.Columns(2).Width = 5270
'                    DataGrid1.Columns(0).Caption = "Group Code"
'                    DataGrid1.Columns(1).Caption = "Test Code"
'                    DataGrid1.Columns(2).Caption = "   Name of Test"
'                Else
'                    DataGrid1.Visible = False
'                End If
''===============================================
'    End If
    
'DataGrid2.Columns(0).Width = 3415
'DataGrid2.Columns(1).Width = 5020
'DataGrid3.Columns(0).Width = 3415
'DataGrid3.Columns(1).Width = 5020
'DataGrid4.Columns(0).Width = 3415
'DataGrid4.Columns(1).Width = 5020

       
End Sub
Public Sub Temp_rst1()

'    If temp_open1 = False Then
'        Temp_Table1.Close
'        temp_open1 = True
'    End If
'
'    If temp_open1 = True Then
'        With Temp_Table1
'            .Fields.Append "Test_Name", adVarChar, 500
'            .Fields.Append "Test_Result", adVarChar, 500
'            .LockType = adLockOptimistic
'            .Open
'            temp_open1 = False
'        End With
'            Set DataGrid2.DataSource = Temp_Table1
'            DataGrid2.ReBind
'            DataGrid2.Refresh
'
'    End If
    Set Temp_Table1 = New ADODB.Recordset
    With Temp_Table1
        .Fields.Append "Test_Name", adVarChar, 500
        .Fields.Append "Test_Result", adVarChar, 500
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set DataGrid2.DataSource = Temp_Table1
    
    DataGrid2.Columns(0).DataField = "Test_Name"
    DataGrid2.Columns(1).DataField = "Test_Result"
    DataGrid2.ReBind
    DataGrid2.Refresh
    
    DataGrid2.Columns(0).Width = 3479.811
    DataGrid2.Columns(1).Width = 4949.858
    
End Sub
Public Sub Temp_rst2()

'    If temp_open2 = False Then
'        Temp_Table2.Close
'        temp_open2 = True
'    End If
'
'    If temp_open2 = True Then
'        With Temp_Table2
'            .Fields.Append "Test_Name1", adVarChar, 500
'            .Fields.Append "Test_Result1", adVarChar, 500
'            .LockType = adLockOptimistic
'            .Open
'            temp_open2 = False
'        End With
'            Set DataGrid3.DataSource = Temp_Table2
'            DataGrid3.ReBind
'            DataGrid3.Refresh
'
'    End If

    Set Temp_Table2 = New ADODB.Recordset
    With Temp_Table2
        .Fields.Append "Test_Name1", adVarChar, 500
        .Fields.Append "Test_Result1", adVarChar, 500
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set DataGrid3.DataSource = Temp_Table2
    
    DataGrid3.Columns(0).DataField = "Test_Name"
    DataGrid3.Columns(1).DataField = "Test_Result"
    DataGrid3.ReBind
    DataGrid3.Refresh
    
    DataGrid3.Columns(0).Width = 3465.071
    DataGrid3.Columns(1).Width = 4995.213



End Sub
Public Sub Temp_rst3()

'    If temp_open3 = False Then
'        Temp_Table3.Close
'        temp_open3 = True
'    End If
'
'    If temp_open3 = True Then
'        With Temp_Table3
'            .Fields.Append "Test_Name2", adVarChar, 500
'            .Fields.Append "Test_Result2", adVarChar, 500
'            .LockType = adLockOptimistic
'            .Open
'            temp_open3 = False
'        End With
'            Set DataGrid4.DataSource = Temp_Table3
'            DataGrid4.ReBind
'            DataGrid4.Refresh
'
'    End If

    Set Temp_Table3 = New ADODB.Recordset
    With Temp_Table3
        .Fields.Append "Test_Name2", adVarChar, 500
        .Fields.Append "Test_Result2", adVarChar, 500
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set DataGrid4.DataSource = Temp_Table3
    
    DataGrid4.Columns(0).DataField = "Test_Name"
    DataGrid4.Columns(1).DataField = "Test_Result"
    DataGrid4.ReBind
    DataGrid4.Refresh
    
    DataGrid4.Columns(0).Width = 3465.071
    DataGrid4.Columns(1).Width = 4995.213

'Me.txtUsed_tech
End Sub
Private Sub InsReport_All_TempRst1()
    
'    Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
'    While Temp_Table1.EOF = False
    
          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
            "','" + txtM_Code + _
            "','" + txtS_Code + _
            "','" + txtSpecimen + _
            "','" + txtN_Exam + _
            "','" + Trim(ComTest_Name.text) + _
            "','" + txtTest_Result.text + _
            "','" + txtUnit.text + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + Trim(txtSN.text) + _
            "','" + txtNote + _
            "','" + u_id + _
            "','" + Format(Dt, "yyyy-mm-dd") + _
            "','" + "3" + _
            "','" + txtPat_ID1 + "'"
            cmd.Execute
'            Temp_Table1.MoveNext
'    Wend
    con.Close
End Sub
'Private Sub InsReport_All_TempRst2()
'
'    Temp_Table2.MoveFirst
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'    While Temp_Table2.EOF = False
'
'          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
'            "','" + txtM_Code + _
'            "','" + txtS_Code + _
'            "','" + txtSpecimen + _
'            "','" + txtN_Exam + _
'            "','" + "" + _
'            "','" + txtChm_Exam + _
'            "','" + Temp_Table2!Test_Name1 + _
'            "','" + Temp_Table2!test_result1 + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + txtNote + _
'            "','" + u_id + _
'            "','" + Format(Dt, "yyyy-mm-dd") + _
'            "','" + "4" + "'"
'            cmd.Execute
'            Temp_Table2.MoveNext
'    Wend
'    con.Close
'End Sub
'Private Sub InsReport_All_TempRst3()
'    Temp_Table3.MoveFirst
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'    While Temp_Table3.EOF = False
'
'          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
'            "','" + txtM_Code + _
'            "','" + txtS_Code + _
'            "','" + txtSpecimen + _
'            "','" + txtN_Exam + _
'            "','" + "" + _
'            "','" + txtMicro_Exam + _
'            "','" + Temp_Table3!Test_Name2 + _
'            "','" + Temp_Table3!test_result2 + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + txtNote + _
'            "','" + u_id + _
'            "','" + Format(Dt, "yyyy-mm-dd") + _
'            "','" + "5" + "'"
'            cmd.Execute
'            Temp_Table3.MoveNext
'    Wend
'    con.Close
'End Sub
'Private Sub UpdReport_All_TempRst1()
'
'    Temp_Table1.MoveFirst
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'    While Temp_Table1.EOF = False
'
'          cmd.CommandText = "exec pro_Report_All 'U','" + Trim(txtPat_ID) + _
'            "','" + txtM_Code + _
'            "','" + txtS_Code + _
'            "','" + txtSpecimen + _
'            "','" + txtSpecimen_dt_Time + _
'            "','" + txtN_Exam + _
'            "','" + "" + _
'            "','" + Temp_Table1!Test_Name + _
'            "','" + Temp_Table1!Test_Result + _
'            "','" + Temp_Table1!Unit + _
'            "','" + Temp_Table1!Ref_Range + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + "" + _
'            "','" + txtNote + _
'            "','" + u_id + _
'            "','" + Format(Dt, "yyyy-mm-dd") + _
'            "','" + "1" + "'"
'            cmd.Execute
'            Temp_Table1.MoveNext
'    Wend
'    con.Close
'End Sub
'Private Sub DelReport_All_TempRst1()
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.Text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "3" + "'"
'            cmd.Execute
'    con.Close
'End Sub
'Private Sub DelReport_All_TempRst2()
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.Text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "4" + "'"
'            cmd.Execute
'    con.Close
'End Sub
'Private Sub DelReport_All_TempRst3()
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.Text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "5" + "'"
'            cmd.Execute
'    con.Close
'End Sub


Private Sub Del_All_Report_All_TempRst1()
   
'    Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
'    While Temp_Table1.EOF = False
            cmd.CommandText = "exec Report_All_Delete2 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
            cmd.Execute
'            Temp_Table1.MoveNext
'    Wend
    con.Close
End Sub
'Private Sub txtSpecimen_dt_Time_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    SendKeys Chr(9)
'    End If
'End Sub
Private Sub txtSpecimen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub txttest_result_GotFocus()
    'GetResult

End Sub

Private Sub txtTest_Result_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub

Private Sub txtTest_Result1_GotFocus()
'    GetResult1
End Sub

Private Sub txtTest_Result1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub

Private Sub txtTest_Result1_LostFocus()
If Trim(ComTest_Name1) = "" Then Exit Sub
'----------------check--------
Dim Check As Integer
Check = 0
If Temp_Table2.RecordCount > 0 Then
    Temp_Table2.MoveFirst
    
        While Temp_Table2.EOF = False
                
            If Temp_Table2!Test_Name1 = ComTest_Name1 Then
                Check = 1
            End If
    Temp_Table2.MoveNext
        Wend
    If Check = 1 Then
        MsgBox "This Test Name already exists"
        ComTest_Name1 = ""
        txtTest_Result1 = ""
        Check = 0
        ComTest_Name1.SetFocus
        Exit Sub
    End If
'    Temp_Table.MoveFirst
End If

'--------------end check-----

'+++to insert into TEMPORARY RECORDSET "Temp_rst2"++++
        Temp_Table2.AddNew
        Temp_Table2!Test_Name1 = ComTest_Name1
        Temp_Table2!test_result1 = txtTest_Result1
        DataGrid3.Refresh
'+++++++++++++++++++++++++++++++++++++++
'    DataGrid3.Columns(0).Width = 2000
'    DataGrid3.Columns(1).Width = 1000
    
    ComTest_Name1.SetFocus
    
ComTest_Name1 = ""
txtTest_Result1 = ""


DataGrid3.Columns(0).Width = 3415
DataGrid3.Columns(1).Width = 5020
End Sub
Private Sub txtTest_Result2_GotFocus()
'    GetResult2
End Sub
Private Sub txtTest_Result2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub

Private Sub txtTest_Result2_LostFocus()
    If Trim(ComTest_Name2) = "" Then Exit Sub
'----------------check--------
Dim Check As Integer
Check = 0
If Temp_Table3.RecordCount > 0 Then
    Temp_Table3.MoveFirst
    
        While Temp_Table3.EOF = False
                
            If Temp_Table3!Test_Name2 = ComTest_Name2 Then
                Check = 1
            End If
    Temp_Table3.MoveNext
        Wend
    If Check = 1 Then
        MsgBox "This Test Name already exists"
        
        ComTest_Name2 = ""
        txtTest_Result2 = ""
        
        Check = 0
        ComTest_Name2.SetFocus
        Exit Sub
    End If
'    Temp_Table.MoveFirst
End If

'--------------end check-----

'+++to insert into TEMPORARY RECORDSET "Temp_rst2"++++
        Temp_Table3.AddNew
        Temp_Table3!Test_Name2 = ComTest_Name2
        Temp_Table3!test_result2 = txtTest_Result2
        DataGrid4.Refresh
'+++++++++++++++++++++++++++++++++++++++

'+++to insert into TEMPORARY RECORDSET "Temp_rst2"++++
'        Temp_Table3.AddNew
'        Temp_Table3!Test_Name2 = ComTest_Name2
'        Temp_Table3!test_result2 = txtTest_Result2
'        DataGrid4.Refresh
'+++++++++++++++++++++++++++++++++++++++
'    DataGrid4.Columns(0).Width = 2000
'    DataGrid4.Columns(1).Width = 1000
    
    ComTest_Name2.SetFocus
    
ComTest_Name2 = ""
txtTest_Result2 = ""


DataGrid4.Columns(0).Width = 3415
DataGrid4.Columns(1).Width = 5020

End Sub

'Private Sub txtUsed_tech_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'     SendKeys Chr(9)
'    End If
'End Sub
Private Sub Clearscreen()
    'txtPat_ID = ""
    'txtM_Code = ""
    txtTest_Result.text = ""
    txtUnit.text = ""
    txtS_Code = ""
    txtS_Name = ""
    txtSpecimen = ""
'    txtUsed_tech = ""
'    txtN_Exam = ""
    txtNote = ""
    Dt.value = Date
    'Temp_rst1
    'Temp_rst2
    'Temp_rst3
    txtPat_ID1.SetFocus
End Sub
Private Sub GetTestName()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select1 1,'" + "10" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
       Do Until Adodc7.Recordset.EOF
          ComTest_Name.AddItem Adodc7.Recordset!Test_Name
       Adodc7.Recordset.MoveNext
       Loop
    End If
End Sub
Private Sub GetResult()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 1,'" + Trim(ComTest_Name.text) + "','" + "10" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    txtTest_Result = Adodc7.Recordset!Test_result
    txtUnit.text = Adodc7.Recordset!unit
   ' txtRef_Range = Adodc7.Recordset!ref_range
    End If
End Sub
'Private Sub GetUnit()
'  Adodc7.connectionstring = strcn.Connection
'  Adodc7.RecordSource = "exec test_result_select 1,'" + Trim(ComTest_Name.Text) + "','" + "10" + "'"
'  Adodc7.Refresh
'
'    If Adodc7.Recordset.RecordCount > 0 Then
'    txtTest_Result = Adodc7.Recordset!unit
'    txtUnit = Adodc7.Recordset!unit
'   ' txtRef_Range = Adodc7.Recordset!ref_range
'    End If
'End Sub
'
'Private Sub GetTestName1()
'  Adodc7.connectionstring = strcn.Connection
'  Adodc7.RecordSource = "exec test_result_select1 1,'" + "10B" + "'"
'  Adodc7.Refresh
'
'    If Adodc7.Recordset.RecordCount > 0 Then
'       Do Until Adodc7.Recordset.EOF
'          ComTest_Name1.AddItem Adodc7.Recordset!Test_Name
'       Adodc7.Recordset.MoveNext
'       Loop
'    End If
'End Sub
'Private Sub GetResult1()
'  Adodc7.connectionstring = strcn.Connection
'  Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Name1 + "','" + "10B" + "'"
'  Adodc7.Refresh
'
'    If Adodc7.Recordset.RecordCount > 0 Then
'    txtTest_Result1 = Adodc7.Recordset!test_result
'    txtUnit1 = Adodc7.Recordset!unit
'   ' txtRef_Range = Adodc7.Recordset!ref_range
'    End If
'End Sub
'Private Sub GetTestName2()
'  Adodc7.connectionstring = strcn.Connection
'  Adodc7.RecordSource = "exec test_result_select1 1,'" + "10C" + "'"
'  Adodc7.Refresh
'
'    If Adodc7.Recordset.RecordCount > 0 Then
'       Do Until Adodc7.Recordset.EOF
'          ComTest_Name2.AddItem Adodc7.Recordset!Test_Name
'       Adodc7.Recordset.MoveNext
'       Loop
'    End If
'End Sub
'Private Sub GetResult2()
'  Adodc7.connectionstring = strcn.Connection
'  Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Name2 + "','" + "10C" + "'"
'  Adodc7.Refresh
'
'    If Adodc7.Recordset.RecordCount > 0 Then
'    txtTest_Result2 = Adodc7.Recordset!test_result
'    'txtUnit2 = Adodc7.Recordset!unit
'   ' txtRef_Range = Adodc7.Recordset!ref_range
'    End If
'End Sub




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


Private Sub BindTestName()
    On Error GoTo err_loop
    ComTest_Name.Clear
       con.connectionstring = strcn.Connection
       con.Open
       rs.Open "exec GetTestName '" & Trim(txtM_Code.text) & "','" & Trim(txtPat_ID.text) & "'", con

       If rs.EOF = False Then
          Do Until rs.EOF
            ComTest_Name.AddItem rs!Test_Name
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
  Adodc7.RecordSource = "exec test_Result_Select8 '" & txtPat_ID & "','" & ComTest_Name.text & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    
    'ComTest_Name.AddItem Adodc7.Recordset!test_result
    txtTest_Result = Adodc7.Recordset!Test_result
    txtUnit.text = Adodc7.Recordset!unit
    'txtNormal_Value = Adodc7.Recordset!ref_range
    
    Adodc7.Recordset.MoveNext
    Loop

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

Private Sub GetSpecimen()

  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec Test_Result_Select11 '" & txtPat_ID & "','" & ComTest_Name.text & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    
    'ComTest_Name.AddItem Adodc7.Recordset!test_result
    'txtTest_Result = Adodc7.Recordset!test_result
    'txtUnit.Text = Adodc7.Recordset!unit
    txtSpecimen = Adodc7.Recordset!others
    
    Adodc7.Recordset.MoveNext
    Loop

    End If
End Sub
