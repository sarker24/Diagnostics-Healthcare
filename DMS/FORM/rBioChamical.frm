VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rBio_Chamical 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Lab Report Format [BIO CHEMICAL]"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "rBioChamical.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPName 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2760
      TabIndex        =   38
      Top             =   1200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtTest_Name 
      BorderStyle     =   0  'None
      Height          =   3315
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3750
      Width           =   2820
   End
   Begin VB.TextBox txtPat_ID1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2790
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   870
      Width           =   1260
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "rBioChamical.frx":000C
      Height          =   600
      Left            =   4260
      TabIndex        =   23
      Top             =   2340
      Visible         =   0   'False
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   1058
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
   Begin VB.TextBox txtTest_Title 
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00000000000E+00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   6
      EndProperty
      Height          =   645
      Left            =   2790
      TabIndex        =   35
      Top             =   2310
      Width           =   6810
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
      Left            =   6255
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10140
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   210
      Top             =   2925
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
   Begin VB.TextBox txtNote 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1395
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   9480
      Width           =   9480
   End
   Begin VB.TextBox txtUnit 
      BorderStyle     =   0  'None
      Height          =   3315
      Left            =   7575
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3735
      Width           =   2985
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
      Left            =   5205
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10140
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
      Left            =   8355
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   10140
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
      Left            =   10455
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10140
      Width           =   1050
   End
   Begin VB.TextBox txtTest_Result 
      BorderStyle     =   0  'None
      Height          =   3315
      Left            =   4275
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3735
      Width           =   3270
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
      Left            =   9405
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   10140
      Width           =   1050
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
      Left            =   7305
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   10140
      Width           =   1050
   End
   Begin VB.ComboBox ComTest_Name 
      DataSource      =   "Adodc7"
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   3420
      Width           =   2805
   End
   Begin VB.CommandButton cmdDelete_TempTable1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "D E L E T E"
      Height          =   2025
      Left            =   10590
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3735
      Width           =   285
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   195
      Top             =   2910
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   195
      Top             =   2910
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   195
      Top             =   2940
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1995
      Left            =   1425
      TabIndex        =   14
      Top             =   7155
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   3519
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
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
      ColumnCount     =   4
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
            ColumnWidth     =   3284.788
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3284.788
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1635.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   195
      Top             =   2940
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
      Left            =   195
      Top             =   2940
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   195
      Top             =   2940
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
      Left            =   9660
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   810
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   195
      Top             =   2940
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   195
      Top             =   2940
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
   Begin VB.TextBox txtN_Exam 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2790
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "As Under"
      Top             =   2010
      Width           =   6810
   End
   Begin VB.TextBox txtS_Name 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1575
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox txtS_Code 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   9630
      TabIndex        =   3
      Top             =   1545
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtM_Code 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   870
      Width           =   345
   End
   Begin VB.TextBox txtPat_ID 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2790
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   870
      Visible         =   0   'False
      Width           =   1290
   End
   Begin MSComCtl2.DTPicker Dt 
      Height          =   285
      Left            =   8400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   810
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      Format          =   65273857
      CurrentDate     =   37114
   End
   Begin VB.TextBox txtSpecimen 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2790
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Blood"
      Top             =   1605
      Width           =   3390
   End
   Begin VB.TextBox txtSN 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1380
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   10005
      Visible         =   0   'False
      Width           =   9540
   End
   Begin VB.Label PName 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Patient Name"
      Height          =   255
      Left            =   960
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Note"
      Height          =   195
      Left            =   1380
      TabIndex        =   36
      Top             =   9750
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Specimen"
      Height          =   195
      Left            =   960
      TabIndex        =   34
      Top             =   1620
      Width           =   1695
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Normal Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7785
      TabIndex        =   33
      Top             =   3090
      Width           =   1230
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bio Chemical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      TabIndex        =   32
      Top             =   165
      Width           =   1860
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   6285
      TabIndex        =   31
      Top             =   1590
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Code"
      Height          =   195
      Left            =   8820
      TabIndex        =   30
      Top             =   1590
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature of Examination"
      Height          =   195
      Left            =   960
      TabIndex        =   29
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      Height          =   195
      Left            =   960
      TabIndex        =   28
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   7845
      TabIndex        =   27
      Top             =   840
      Width           =   345
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1515
      TabIndex        =   26
      Top             =   3060
      Width           =   1155
   End
   Begin VB.Label lblOverflow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5115
      TabIndex        =   25
      Top             =   3060
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      Height          =   240
      Left            =   1395
      TabIndex        =   24
      Top             =   9195
      Width           =   345
   End
End
Attribute VB_Name = "rBio_Chamical"
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
'Dim IntPat_ID As Integer
Dim IntPat_ID As Double

Private Sub cmdClear_Click()
    ComTest_Name.Clear
    Temp_rst1
    txtPat_ID = ""
    txtPat_ID1 = ""
    txtS_Code = ""
    txtS_Name = ""
    Dt.value = Now
    txtNote = ""
    
    txtRef_Range = ""
    txtTest_Result = ""
    txtUnit = ""
    txtSpecimen = "Blood"
    txtN_Exam = "As Under"
    'txtUsed_tech = ""
    txtSN = ""
    StrSub_Code = "0"
    txtTest_Title = ""
    DataGrid1.Visible = False
    GetTestName
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
    Temp_rst1
    txtPat_ID1.SetFocus
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
        DelReport_All_TempRst1
        Temp_Table1.Delete
        ComTest_Name = ""
        txtTest_Result = ""
        txtUnit = ""
        txtRef_Range = ""
        End If
        
    End If
End Sub

Private Sub CmdPreview_Click()
    CRViewer1_MODE = 11
    Viewer.Show vbModal
End Sub

Private Sub cmdPrint_Click()
    '==========direct print==========================
            
            Dim Report11 As New BioChamical
            Dim StrPat_ID As String
            
            Dim strM_Code As String
            Dim strS_Code As String
            
            StrPat_ID = rBio_Chamical.txtPat_ID
            StrPat_ID_R = StrPat_ID
            
            strM_Code = rBio_Chamical.txtM_Code
            strS_Code = rBio_Chamical.txtS_Code
            
            
            '--------------------------------------------------------------------
            Report11.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report11.FormulaFields.Item(2).text = Chr(34) & "Patient ID" & Chr(34)
            Report11.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report11.FormulaFields.Item(4).text = Chr(34) & "Delivered Date" & Chr(34)
            Report11.FormulaFields.Item(5).text = Chr(34) & "Patient Name" & Chr(34)
            Report11.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report11.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report11.FormulaFields.Item(8).text = Chr(34) & "Refd. by" & Chr(34)
            '--------------------------------------------------------------------
            Report11.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report11.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report11.FormulaFields.Item(11).text = Chr(34) & "Biochemical Report" & Chr(34)
            Report11.FormulaFields.Item(12).text = Chr(34) & "Name of Tests" & Chr(34)
            Report11.FormulaFields.Item(13).text = Chr(34) & "Results" & Chr(34)
            Report11.FormulaFields.Item(14).text = Chr(34) & "Unit" & Chr(34)
            Report11.FormulaFields.Item(15).text = Chr(34) & "Normal Values" & Chr(34)
            Report11.FormulaFields.Item(16).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report11.Text1.SetText StDoc_Name
            
            
            Report11.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report11.Database.SetDataSource rs
            
            Report11.PrintOut (False)
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
    
''\\\\\\\\\\INSERT and UPDATE from Temp_rst2\\\\\\\\\\\\\
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "Report_All_SELECT3 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        Del_All_Report_All_TempRst1
        InsReport_All_TempRst1
        MsgBox "Updated1"
    Else
        InsReport_All_TempRst1
        MsgBox "Inserted1"
    End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Temp_rst1
    DataGrid1.Visible = False
    cmdPrint.SetFocus
    
End Sub
Private Sub cmdShow_Click()
If cmdSave.Enabled = False Then Exit Sub

        If txtPat_ID1 = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID1 = ""
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
        txtPat_ID1 = ""
        txtPat_ID = ""
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

Private Sub ComTest_Name_GotFocus()
    'ComTest_Name = ""
    'GetTestName
End Sub

Private Sub ComTest_Name_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub

Private Sub ComTest_Name_LostFocus()
On Error Resume Next
    If Trim(ComTest_Name) = "" Then
        cmdSave.SetFocus
        Exit Sub
    End If
    GetUsed_Tech
    GetNote
End Sub

Private Sub DataGrid1_Click()
    txtS_Code.text = DataGrid1.Columns(1)
    StrSub_Code = txtS_Code.text
    'StrSub_Code = DataGrid2.Columns(1).value
    GetTestName
End Sub

Private Sub DataGrid2_DblClick()
On Error Resume Next
    ComTest_Name.text = DataGrid2.Columns(0)
    txtTest_Result.text = DataGrid2.Columns(1)
    txtUnit.text = DataGrid2.Columns(2)
    'txtRef_Range = DataGrid2.Columns(3)
End Sub

Private Sub Form_Click()
    DataGrid1.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    If rBio_Chamical.DataGrid1.Visible = True Then
        rBio_Chamical.DataGrid1.Visible = False
    Else
        Unload Me
    End If
    End If
End Sub

Private Sub Form_Load()
    
    
    Adodc8.connectionstring = strcn.Connection
'    Adodc8.RecordSource = "exec m_name_select 2,'" + "BIOCHEMICAL EXAMINATION" + "'"
    Adodc8.RecordSource = "exec m_name_select 2,'" & "BIO" & "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
        txtM_Code = "02"
    Else
        MsgBox "Inserted incurrect head name, first you have to insert currect name from TEST INFORMATION form then open this screen again"
        txtPat_ID.Enabled = False
        cmdSave.Enabled = False
    End If


    Dt.value = Date
    Temp_rst1
'------select data from test_result------
'     GetTestName
'------end------------------------------

DataGrid2.Columns(0).Width = 2800.142
DataGrid2.Columns(1).Width = 3300.095
DataGrid2.Columns(2).Width = 3000
DataGrid2.Columns(3).Width = 0
    StrScreenName = "Bio"
    Flush_Font_Type
    'MsgBox IntFont
    
    
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
Private Sub txtPat_ID_LostFocus()
'Pat_Paid

'If Len(Trim(txtPat_ID.Text)) = 0 Then Exit Sub
'
'
'Adodc5.connectionstring = strcn.Connection
'Adodc5.RecordSource = "exec Pro_FLUSH 6," & txtPat_ID & ""
'Adodc5.Refresh
'If Adodc5.Recordset.RecordCount > 0 Then
'
'Else
'    '===for show data in Datagrid1=============
'                Adodc1.connectionstring = strcn.Connection
'                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" & txtM_Code & "','" & txtPat_ID & "'"
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
    StrSub_Code = txtS_Code.text
    txtS_Name.text = DataGrid1.Columns(2)
    txtSpecimen.SetFocus
    DataGrid1.Visible = False
    
    GetTestName
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
                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" & txtM_Code & "','" & txtPat_ID & "'"
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
                    txtPat_ID1 = ""
                    txtPat_ID1.SetFocus
                    
                End If
        '===============================================
End If

'-->>>>>>FOR SHOW PREVOIUS DATA --->>>>

Temp_rst1
    
'If Len(txtS_Code.Text) = 0 Then Exit Sub
If cmdSave.Enabled = False Then Exit Sub

If Trim(txtPat_ID1) = "" Then
    MsgBox "Patient ID mandatory"
    txtPat_ID = ""
    txtPat_ID1.SetFocus
    Exit Sub
End If

'If Len(Trim(txtS_Code)) = 0 Then Exit Sub
         
    Adodc6.connectionstring = strcn.Connection
    Adodc6.RecordSource = "exec Report_All_Select2 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc6.Refresh
    
    If Adodc6.Recordset.RecordCount > 0 Then
        DataGrid1.Visible = False
        txtSpecimen = Adodc6.Recordset!Field1
         txtN_Exam = Adodc6.Recordset!Field2
         txtTest_Title = Adodc6.Recordset!Field3
         txtSN = Adodc6.Recordset!Field14
         txtNote = Adodc6.Recordset!Field15
'         txtUsed_tech = Adodc6.Recordset!Field3
         
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
'                Temp_Table1!ref_range = Temp_Table_Helper1!Field7
                Temp_Table_Helper1.MoveNext
            Wend
        DataGrid2.Refresh
        Temp_Table_Helper1.Close
        con.Close
'/////////end show in Temp_rst1////////////////////////////

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

'---<<<<------------------------------

DataGrid2.Columns(0).Width = 2800.142
DataGrid2.Columns(1).Width = 3300.095
DataGrid2.Columns(2).Width = 3000
DataGrid2.Columns(3).Width = 0

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

Public Sub Temp_rst1()
    
    Set Temp_Table1 = New ADODB.Recordset
    With Temp_Table1
        .Fields.Append "Test_Name", adVarChar, 500
        .Fields.Append "Test_Result", adVarChar, 500
        .Fields.Append "Unit", adVarChar, 500
        '.Fields.Append "Ref_Range", adVarChar, 500
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set DataGrid2.DataSource = Temp_Table1
    
    DataGrid2.Columns(0).DataField = "Test_Name"
    DataGrid2.Columns(1).DataField = "Test_Result"
    DataGrid2.Columns(2).DataField = "Unit"
    'DataGrid2.Columns(3).DataField = "Ref_Range"
    DataGrid2.ReBind
    DataGrid2.Refresh
    
    DataGrid2.Columns(0).Width = 2800.142
    DataGrid2.Columns(1).Width = 3300.095
    DataGrid2.Columns(2).Width = 3000
    DataGrid2.Columns(3).Width = 0


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
            "','" + txtSpecimen + _
            "','" + txtN_Exam + _
            "','" + Trim(txtTest_Title) + _
            "','" + Temp_Table1!Test_Name + _
            "','" + Temp_Table1!Test_result + _
            "','" + Temp_Table1!unit + _
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
            "','" + "" + _
            "','" + txtPat_ID1.text + "'"
            cmd.Execute
            Temp_Table1.MoveNext
    Wend
    con.Close
End Sub

Private Sub DelReport_All_TempRst1()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "" + "'"
            cmd.Execute
    con.Close
End Sub
Private Sub Del_All_Report_All_TempRst1()
   On Error Resume Next
    Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    While Temp_Table1.EOF = False
            cmd.CommandText = "exec Report_All_Delete2 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
            cmd.Execute
            Temp_Table1.MoveNext
    Wend
    con.Close
End Sub
Private Sub txtSpecimen_dt_Time_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub
Private Sub txtSpecimen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub

Private Sub txttest_result_GotFocus()
    GetResult
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
    txtPat_ID = ""
    txtPat_ID1 = ""
    txtSpecimen = "Blood"
    txtN_Exam = "As Under"
    txtNote = ""
    Dt.value = Date
    txtTest_Title = ""
    ComTest_Name.Clear
    
    DataGrid1.Visible = False
End Sub

Private Sub GetTestName()
    'MsgBox StrSub_Code
  Adodc7.connectionstring = strcn.Connection
'  Adodc7.RecordSource = "exec test_result_select1 1,'" & "02" & "'"
  Adodc7.RecordSource = "exec Flush_Test_Result 1,'" & txtM_Code & "','" & txtM_Code & "','" & StrSub_Code & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
       Do Until Adodc7.Recordset.EOF
          ComTest_Name.AddItem Adodc7.Recordset!Test_Name
          'Dim SSTT As String
          'SSTT = Adodc7.Recordset!Test_Name
       Adodc7.Recordset.MoveNext
       Loop
    End If
    'MsgBox SSTT
End Sub
Private Sub GetResult()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 1,'" & ComTest_Name & "','" & txtM_Code & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    txtTest_Result = Adodc7.Recordset!Test_result
    txtUnit = Adodc7.Recordset!unit
    txtRef_Range = Adodc7.Recordset!ref_range
    txtTest_Name.text = ComTest_Name.text
    'txtSpecimen = Adodc7.Recordset!others
    'txtN_Exam = Adodc7.Recordset!others1
    
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

Private Sub txtUnit_LostFocus()

If Trim(ComTest_Name) = "" Then Exit Sub
'----------------check--------
Dim Check As Integer
Check = 0
If Temp_Table1.RecordCount > 0 Then
    Temp_Table1.MoveFirst
    
        While Temp_Table1.EOF = False
                
            If Temp_Table1!Test_Name = ComTest_Name Then
                Check = 1
            End If
    Temp_Table1.MoveNext
        Wend
    If Check = 1 Then
        MsgBox "This Test Name already exists"
        Check = 0
        ComTest_Name.SetFocus
        Exit Sub
    End If
'    Temp_Table.MoveFirst
End If

'--------------end check-----

'+++to insert into TEMPORARY RECORDSET "Temp_rst1"++++
        Temp_Table1.AddNew
        Temp_Table1!Test_Name = ComTest_Name
        Temp_Table1!Test_result = txtTest_Result
        Temp_Table1!unit = txtUnit
'        Temp_Table1!ref_range = txtRef_Range
        DataGrid2.Refresh
'+++++++++++++++++++++++++++++++++++++++
 '   DataGrid2.Columns(0).Width = 2000
    DataGrid2.Columns(0).Width = 1000
    
    ComTest_Name.SetFocus
    
ComTest_Name = ""
txtTest_Result = ""
txtUnit = ""
'txtRef_Range = ""

DataGrid2.Columns(0).Width = 2970.142
DataGrid2.Columns(1).Width = 3300.095
DataGrid2.Columns(2).Width = 1769.953
'DataGrid2.Columns(3).Width = 1785.26


End Sub
Private Sub GetUsed_Tech()

  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_Result_Select10 '" & txtPat_ID & "','" & ComTest_Name.text & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    
    'ComTest_Name.AddItem Adodc7.Recordset!test_result
    'txtTest_Result = Adodc7.Recordset!test_result
    'txtUnit.Text = Adodc7.Recordset!unit
    txtTest_Title.text = Adodc7.Recordset!ref_range
    
    Adodc7.Recordset.MoveNext
    Loop

    End If
End Sub



Private Sub GetNote()

  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_Result_Select11 '" & txtPat_ID & "','" & ComTest_Name.text & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    
    'ComTest_Name.AddItem Adodc7.Recordset!test_result
    'txtTest_Result = Adodc7.Recordset!test_result
    'txtUnit.Text = Adodc7.Recordset!unit
    txtNote.text = Adodc7.Recordset!others
    
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

