VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rStool 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Diagnostic management system"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   DrawWidth       =   2
   Icon            =   "rStool.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPat_ID1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   480
      Width           =   1260
   End
   Begin VB.TextBox txtSN 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1770
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   7800
      Width           =   7950
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "rStool.frx":000C
      Height          =   1020
      Left            =   4320
      TabIndex        =   20
      Top             =   420
      Visible         =   0   'False
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   1799
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
   Begin VB.TextBox txtTest_Name3 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   3510
   End
   Begin VB.TextBox txtTest_Name2 
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2940
      Width           =   3510
   End
   Begin VB.TextBox txtTest_Name1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   3510
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
      Left            =   4830
      TabIndex        =   15
      Top             =   8130
      Width           =   1050
   End
   Begin VB.TextBox txtPat_ID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2775
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtM_Code 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   345
   End
   Begin VB.TextBox txtS_Code 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2790
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   765
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
      Left            =   5880
      TabIndex        =   16
      Top             =   8130
      Width           =   1050
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
      Left            =   7980
      TabIndex        =   18
      Top             =   8130
      Width           =   1050
   End
   Begin VB.TextBox txtTest_Result1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   5370
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   3900
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
      Left            =   9030
      TabIndex        =   19
      Top             =   8130
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
      Left            =   6930
      TabIndex        =   17
      Top             =   8130
      Width           =   1050
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
      Left            =   3780
      TabIndex        =   14
      Top             =   8130
      Width           =   1050
   End
   Begin VB.TextBox txtNote 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1770
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   7500
      Width           =   7950
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
      Height          =   240
      Left            =   1800
      TabIndex        =   23
      Text            =   "PHYSICAL EXAMINATION"
      Top             =   1170
      Width           =   3405
   End
   Begin VB.TextBox txtTest_Result2 
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   5370
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2940
      Width           =   3900
   End
   Begin VB.TextBox txtChm_Exam 
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
      Height          =   240
      Left            =   1800
      TabIndex        =   22
      Text            =   "CHEMICAL EXAMINATION"
      Top             =   2670
      Width           =   3435
   End
   Begin VB.TextBox txtTest_Result3 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   5370
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4320
      Width           =   3900
   End
   Begin VB.TextBox txtMicro_Exam 
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
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Text            =   "MICROSCOPIC EXAMINATION"
      Top             =   4035
      Width           =   3405
   End
   Begin VB.TextBox txtS_Name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   750
      Visible         =   0   'False
      Width           =   5115
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
      Left            =   9390
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Width           =   1050
   End
   Begin MSComCtl2.DTPicker Dt 
      Height          =   285
      Left            =   8115
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      Format          =   66781185
      CurrentDate     =   37114
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   7710
      Top             =   150
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
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
      Left            =   9810
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
      Left            =   9810
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
      Left            =   9840
      Top             =   120
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
      Left            =   9780
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
      Left            =   9780
      Top             =   60
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
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
      Left            =   9840
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
      Left            =   9840
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
      Left            =   9840
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
      Left            =   9840
      Top             =   60
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
      Left            =   9840
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Note"
      Height          =   195
      Left            =   720
      TabIndex        =   31
      Top             =   7725
      Width           =   915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      Height          =   240
      Left            =   750
      TabIndex        =   29
      Top             =   7455
      Width           =   345
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   7560
      TabIndex        =   28
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      Height          =   195
      Left            =   960
      TabIndex        =   27
      Top             =   510
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Code"
      Height          =   195
      Left            =   960
      TabIndex        =   26
      Top             =   795
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   3750
      TabIndex        =   25
      Top             =   780
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOOL"
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
      Left            =   5040
      TabIndex        =   24
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "rStool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Strpat_id1 As String
Dim StrRow_Count As String
Dim StrPat_Type As String
Dim IntPat_ID As Double

Private Sub cmdClear_Click()
    txtS_Code.text = ""
    txtTest_Name1.text = ""
    txtTest_Result1.text = ""
    txtTest_Name2.text = ""
    txtTest_Result2.text = ""
    txtTest_Name3.text = ""
    txtTest_Result3.text = ""
    txtNote.text = ""
    txtSN.text = ""
    GetTestName1
    GetTestName2
    GetTestName3
    txtPat_ID = ""
    txtPat_ID1 = ""
    DataGrid1.Visible = False
    txtPat_ID1.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    If txtPat_ID.text = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub

    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
    'Del_All_Report_All_TempRst1
    Del_Report
    Clearscreen
    
    GetTestName1
    GetTestName2
    GetTestName3
    txtPat_ID = ""
    txtPat_ID1 = ""
    DataGrid1.Visible = False
    txtPat_ID1.SetFocus
    End If
End Sub

Private Sub CmdPreview_Click()
If txtPat_ID1 = "" Then Exit Sub

    CRViewer1_MODE = 12
    Viewer.Show vbModal
    
End Sub

Private Sub cmdPrint_Click()
    '==========direct print==========================
            
            Dim Report12 As New Stool
            Dim StrPat_ID As String
            
            Dim strM_Code As String
            Dim strS_Code As String
            
            StrPat_ID = rStool.txtPat_ID
            StrPat_ID_R = StrPat_ID
            
            strM_Code = rStool.txtM_Code
            strS_Code = rStool.txtS_Code
            
            '--------------------------------------------------------------------
            Report12.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report12.FormulaFields.Item(2).text = Chr(34) & "Patient ID" & Chr(34)
            Report12.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report12.FormulaFields.Item(4).text = Chr(34) & "Delivered Date" & Chr(34)
            Report12.FormulaFields.Item(5).text = Chr(34) & "Patient Name" & Chr(34)
            Report12.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report12.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report12.FormulaFields.Item(8).text = Chr(34) & "Refd. by" & Chr(34)
            '--------------------------------------------------------------------
            Report12.FormulaFields.Item(9).text = Chr(34) & "STOOL EXAMINATION REPORT" & Chr(34)
            Report12.FormulaFields.Item(10).text = Chr(34) & "Checked By" & Chr(34)
            
            Call Flush_Doc_Name
            Report12.Text1.SetText StDoc_Name
  
            
            Report12.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report12.Database.SetDataSource rs
            
            Report12.PrintOut (False)
            rs.Close
            Call cmdClear_Click
            txtPat_ID1.SetFocus
    '====================================
End Sub

Private Sub cmdSave_Click()
'-----validation check---------------------
    If Trim(txtPat_ID1) = "" Then
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
        InsReport_All_TempRst2
        InsReport_All_TempRst3
        MsgBox "Updated"
        
    Else
        InsReport_All_TempRst1
        InsReport_All_TempRst2
        InsReport_All_TempRst3
        MsgBox "Inserted"
    End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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

Private Sub Form_Click()
    DataGrid1.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If DataGrid1.Visible = False Then
            Unload Me
        End If
    
        If DataGrid1.Visible = True Then
            DataGrid1.Visible = False
        End If
        
    End If
End Sub

Private Sub Form_Load()
    
    Adodc8.connectionstring = strcn.Connection
    Adodc8.RecordSource = "exec m_name_select 2,'" + "STOOL" + "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
        txtM_Code = Adodc8.Recordset!m_code
    Else
        MsgBox "Inserted incurrect head name, first you have to insert currect name from TEST INFORMATION form then open this screen again"
        txtPat_ID.Enabled = False
        cmdSave.Enabled = False
    End If


    Dt.value = Date
    GetTestName1
    GetTestName2
    GetTestName3
    
StrScreenName = "Stool"
Flush_Font_Type
    
End Sub
Private Sub txtPat_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub txtPat_ID_LostFocus()
'Pat_Paid
'If Len(Trim(txtPat_ID.Text)) = 0 Then Exit Sub
'
'Adodc5.connectionstring = strcn.Connection
'Adodc5.RecordSource = "exec Pro_FLUSH 6," & txtPat_ID & ""
'Adodc5.Refresh
'If Adodc5.Recordset.RecordCount > 0 Then
'
'Else
'    '===for show data in Datagrid1=============
'                Adodc1.connectionstring = strcn.Connection
'                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
'                Adodc1.Refresh
'                If Adodc1.Recordset.RecordCount > 0 Then
'                    DataGrid1.Visible = True
'                    DataGrid1.Columns(2).Width = 5270
'                    DataGrid1.Columns(0).Caption = "Group Code"
'                    DataGrid1.Columns(1).Caption = "Test Code"
'                    DataGrid1.Columns(2).Caption = "   Name of Test"
'                Else
'                    DataGrid1.Visible = False
'                    MsgBox "Invalied Patient ID"
'                    txtPat_ID = ""
'                    txtPat_ID.SetFocus
'
'                End If
'        '===============================================
'End If


End Sub
Private Sub DataGrid1_DblClick()
    'txtM_Code.Text = DataGrid1.Columns(0)
    txtS_Code.text = DataGrid1.Columns(1)
    txtS_Name.text = DataGrid1.Columns(2)
    txtTest_Name1.SetFocus
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
                    DataGrid1.Visible = True
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

'------------------------------------------
If cmdSave.Enabled = False Then Exit Sub

If Trim(txtPat_ID1) = "" Then
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
         txtPh_Exam.text = Adodc6.Recordset!Field1
         txtTest_Name1.text = Adodc6.Recordset!Field2
         txtTest_Result1.text = Adodc6.Recordset!Field3
         txtSN.text = Adodc6.Recordset!Field14
         txtNote = Adodc6.Recordset!Field15
         Dt.value = Adodc6.Recordset!Dt


'++++++++++for show feild5 to CHAMICAL EXAMINATION FROM TYPE 4 +++++++++++
    Adodc10.connectionstring = strcn.Connection
    Adodc10.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "4" + "'"
    Adodc10.Refresh
    If Adodc10.Recordset.RecordCount > 0 Then
        txtChm_Exam.text = Adodc10.Recordset!Field1
        txtTest_Name2.text = Adodc10.Recordset!Field2
        txtTest_Result2.text = Adodc10.Recordset!Field3

    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++for show feild5 to MICROSCOPIC EXAMINATION FROM TYPE 5 +++++++++++
    Adodc11.connectionstring = strcn.Connection
    Adodc11.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "5" + "'"
    Adodc11.Refresh
    If Adodc11.Recordset.RecordCount > 0 Then
        txtMicro_Exam.text = Adodc11.Recordset!Field1
        txtTest_Name3.text = Adodc11.Recordset!Field2
        txtTest_Result3.text = Adodc11.Recordset!Field3

    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
         


    Else
    '===for show data in Datagrid1=============
                Adodc1.connectionstring = strcn.Connection
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

If cmdSave.Enabled = False Then Exit Sub

If Trim(txtPat_ID1) = "" Then
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
         txtPh_Exam.text = Adodc6.Recordset!Field1
         txtTest_Name1.text = Adodc6.Recordset!Field2
         txtTest_Result1.text = Adodc6.Recordset!Field3
         txtSN.text = Adodc6.Recordset!Field14
         txtNote = Adodc6.Recordset!Field15
         Dt.value = Adodc6.Recordset!Dt


'++++++++++for show feild5 to CHAMICAL EXAMINATION FROM TYPE 4 +++++++++++
    Adodc10.connectionstring = strcn.Connection
    Adodc10.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "4" + "'"
    Adodc10.Refresh
    If Adodc10.Recordset.RecordCount > 0 Then
        txtChm_Exam.text = Adodc10.Recordset!Field1
        txtTest_Name2.text = Adodc10.Recordset!Field2
        txtTest_Result2.text = Adodc10.Recordset!Field3

    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++for show feild5 to MICROSCOPIC EXAMINATION FROM TYPE 5 +++++++++++
    Adodc11.connectionstring = strcn.Connection
    Adodc11.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "5" + "'"
    Adodc11.Refresh
    If Adodc11.Recordset.RecordCount > 0 Then
        txtMicro_Exam.text = Adodc11.Recordset!Field1
        txtTest_Name3.text = Adodc11.Recordset!Field2
        txtTest_Result3.text = Adodc11.Recordset!Field3

    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
         


    Else
    '===for show data in Datagrid1=============
                Adodc1.connectionstring = strcn.Connection
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

       
End Sub
Private Sub InsReport_All_TempRst1()
    
    'Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    'While Temp_Table1.EOF = False
    
          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
            "','" + txtM_Code + _
            "','" + txtS_Code + _
            "','" + Trim(txtPh_Exam.text) + _
            "','" + Trim(txtTest_Name1.text) + _
            "','" + Trim(txtTest_Result1.text) + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + Trim(txtSN) + _
            "','" + txtNote + _
            "','" + u_id + _
            "','" + Format(Dt, "yyyy-mm-dd") + _
            "','" + "3" + _
            "','" + txtPat_ID1.text + "'"
            cmd.Execute
     '       Temp_Table1.MoveNext
    'Wend
    con.Close
End Sub
Private Sub InsReport_All_TempRst2()
    'Temp_Table2.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    'While Temp_Table2.EOF = False
    
          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
            "','" + txtM_Code + _
            "','" + txtS_Code + _
            "','" + Trim(txtChm_Exam) + _
            "','" + txtTest_Name2.text + _
            "','" + txtTest_Result2.text + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + Trim(txtSN) + _
            "','" + txtNote + _
            "','" + u_id + _
            "','" + Format(Dt, "yyyy-mm-dd") + _
            "','" + "4" + _
            "','" + txtPat_ID1.text + "'"
            cmd.Execute
            'Temp_Table2.MoveNext
    'Wend
    con.Close
End Sub
Private Sub InsReport_All_TempRst3()
    'Temp_Table3.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    'While Temp_Table3.EOF = False
    
          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
            "','" + txtM_Code + _
            "','" + txtS_Code + _
            "','" + Trim(txtMicro_Exam.text) + _
            "','" + txtTest_Name3.text + _
            "','" + txtTest_Result3.text + _
            "','" + "" + _
            "','" + "" + _
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
            "','" + "5" + _
            "','" + txtPat_ID1.text + "'"
            cmd.Execute
'            Temp_Table3.MoveNext
'    Wend
    con.Close
End Sub
Private Sub DelReport_All_TempRst1()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "3" + "'"
            cmd.Execute
    con.Close
End Sub
Private Sub DelReport_All_TempRst2()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "4" + "'"
            cmd.Execute
    con.Close
End Sub
Private Sub DelReport_All_TempRst3()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "5" + "'"
            cmd.Execute
    con.Close
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

Private Sub txtUsed_tech_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub Clearscreen()
    'txtPat_ID = ""
    'txtM_Code = ""
    txtS_Code = ""
    txtS_Name = ""
    'txtSpecimen = ""
    'txtUsed_tech = ""
    'txtN_Exam = ""
    txtNote = ""
    Dt.value = Date
    
End Sub

Private Sub GetTestName1()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 1,'" + Trim(txtPh_Exam.text) + "','" + "09" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    txtTest_Name1.text = Adodc7.Recordset!Test_result
    txtTest_Result1 = Adodc7.Recordset!unit
    End If
End Sub
Private Sub GetTestName2()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 1,'" + Trim(txtChm_Exam.text) + "','" + "09" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    txtTest_Name2.text = Adodc7.Recordset!Test_result
    txtTest_Result2 = Adodc7.Recordset!unit
    End If
End Sub
Private Sub GetTestName3()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 1,'" + Trim(txtMicro_Exam.text) + "','" + "09" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    txtTest_Name3.text = Adodc7.Recordset!Test_result
    txtTest_Result3 = Adodc7.Recordset!unit
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

Private Sub Del_Report()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Del_Report 1,'" + Trim(txtPat_ID.text) + "'"
            cmd.Execute
    con.Close
End Sub

