VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rUrine1 
   Caption         =   "Prime Diagnostic Ltd."
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   Icon            =   "rUrine1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   3090
      TabIndex        =   42
      Top             =   8010
      Width           =   1050
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "rUrine1.frx":030A
      Height          =   1200
      Left            =   2010
      TabIndex        =   22
      Top             =   1290
      Visible         =   0   'False
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   2117
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
   Begin VB.CommandButton cmdDelete_TempTable1 
      Caption         =   "D ELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   8985
      TabIndex        =   31
      Top             =   2925
      Width           =   315
   End
   Begin VB.TextBox txtSpecimen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2055
      TabIndex        =   6
      Top             =   1305
      Width           =   3390
   End
   Begin VB.TextBox txtSpecimen_dt_Time 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2055
      TabIndex        =   7
      Top             =   1590
      Width           =   3390
   End
   Begin VB.TextBox txtPat_ID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2055
      MaxLength       =   7
      TabIndex        =   0
      Top             =   735
      Width           =   1050
   End
   Begin VB.TextBox txtM_Code 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3525
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   345
   End
   Begin VB.TextBox txtS_Code 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2070
      TabIndex        =   2
      Top             =   1020
      Width           =   765
   End
   Begin VB.ComboBox ComTest_Name 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc7"
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   2940
      Width           =   3405
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
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
      Left            =   8340
      TabIndex        =   21
      Top             =   8010
      Width           =   1050
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
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
      Left            =   5190
      TabIndex        =   18
      Top             =   8010
      Width           =   1050
   End
   Begin VB.TextBox txtTest_Result 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3690
      TabIndex        =   11
      Top             =   2955
      Width           =   4980
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
      Left            =   7290
      TabIndex        =   20
      Top             =   8010
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
      Left            =   6240
      TabIndex        =   19
      Top             =   8010
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
      Left            =   4140
      TabIndex        =   17
      Top             =   8010
      Width           =   1050
   End
   Begin VB.TextBox txtNote 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   7530
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
      Height          =   225
      Left            =   240
      TabIndex        =   28
      Text            =   "PHYSICAL EXAMINATION"
      Top             =   2640
      Width           =   3405
   End
   Begin VB.TextBox txtTest_Result1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3690
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   4485
      Width           =   4980
   End
   Begin VB.ComboBox ComTest_Name1 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc7"
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   4485
      Width           =   3405
   End
   Begin VB.CommandButton cmdDelete_TempTable2 
      Caption         =   "D ELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   8985
      TabIndex        =   27
      Top             =   4485
      Width           =   315
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
      Height          =   225
      Left            =   240
      TabIndex        =   26
      Text            =   "CHEMICAL EXAMINATION"
      Top             =   4230
      Width           =   3435
   End
   Begin VB.CommandButton cmdDelete_TempTable3 
      Caption         =   "D ELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   9000
      TabIndex        =   25
      Top             =   6165
      Width           =   315
   End
   Begin VB.ComboBox ComTest_Name2 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc7"
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   6150
      Width           =   3405
   End
   Begin VB.TextBox txtTest_Result2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3690
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   6180
      Width           =   4980
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
      Left            =   240
      TabIndex        =   24
      Text            =   "MICROSCOPIC EXAMINATION"
      Top             =   5865
      Width           =   3405
   End
   Begin VB.TextBox txtUsed_tech 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2055
      TabIndex        =   9
      Top             =   2160
      Width           =   6615
   End
   Begin VB.TextBox txtN_Exam 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2055
      TabIndex        =   8
      Top             =   1875
      Width           =   6600
   End
   Begin VB.TextBox txtS_Name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3525
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1020
      Width           =   5115
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
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
      Left            =   8850
      TabIndex        =   5
      Top             =   690
      Width           =   1050
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   870
      Left            =   240
      TabIndex        =   23
      Top             =   3270
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   1535
      _Version        =   393216
      BorderStyle     =   0
      ColumnHeaders   =   0   'False
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker Dt 
      Height          =   285
      Left            =   7395
      TabIndex        =   4
      Top             =   690
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24641537
      CurrentDate     =   37114
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   6990
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   930
      Left            =   240
      TabIndex        =   29
      Top             =   4830
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   1640
      _Version        =   393216
      BorderStyle     =   0
      ColumnHeaders   =   0   'False
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   870
      Left            =   240
      TabIndex        =   30
      Top             =   6480
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   1535
      _Version        =   393216
      BorderStyle     =   0
      ColumnHeaders   =   0   'False
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   9120
      Top             =   1050
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
      Left            =   9120
      Top             =   690
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
      Left            =   9120
      Top             =   1410
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
      Left            =   9060
      Top             =   90
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
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
      Left            =   9060
      Top             =   360
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
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
      Left            =   9120
      Top             =   2700
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   9120
      Top             =   3030
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   9120
      Top             =   2400
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
      Left            =   9120
      Top             =   1740
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
      Left            =   9120
      Top             =   2070
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      Height          =   240
      Left            =   210
      TabIndex        =   41
      Top             =   7575
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Used Technology"
      Height          =   195
      Left            =   240
      TabIndex        =   40
      Top             =   2220
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   6840
      TabIndex        =   39
      Top             =   750
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      Height          =   195
      Left            =   240
      TabIndex        =   38
      Top             =   780
      Width           =   705
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature of Examination"
      Height          =   195
      Left            =   240
      TabIndex        =   37
      Top             =   1935
      Width           =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Code"
      Height          =   195
      Left            =   240
      TabIndex        =   36
      Top             =   1065
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   3030
      TabIndex        =   35
      Top             =   1050
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URINE-1"
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
      Left            =   4350
      TabIndex        =   34
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Specimen Date & Time"
      Height          =   195
      Left            =   240
      TabIndex        =   33
      Top             =   1650
      Width           =   1530
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Specimen"
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   1350
      Width           =   750
   End
End
Attribute VB_Name = "rUrine1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp_Table1 As New ADODB.Recordset
Dim Temp_Table_Helper1 As New ADODB.Recordset
Dim Temp_Table2 As New ADODB.Recordset
Dim Temp_Table_Helper2 As New ADODB.Recordset
Dim Temp_Table3 As New ADODB.Recordset
Dim Temp_Table_Helper3 As New ADODB.Recordset
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    If cmdSave.Enabled = False Then Exit Sub

    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
    Del_All_Report_All_TempRst1

    Clearscreen
    End If
End Sub
Private Sub cmdDelete_TempTable1_Click()

    If ComTest_Name = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
    If Temp_Table1.RecordCount <= 0 Then Exit Sub
    
    If Trim(ComTest_Name.Text) = "" Then
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
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_TempTable2_Click()
    If ComTest_Name1 = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
    If Temp_Table2.RecordCount <= 0 Then Exit Sub
    
    If Trim(ComTest_Name.Text) = "" Then
        MsgBox "You didn't select the the Test Name"
        DataGrid3.SetFocus
        Exit Sub
    Else
        Dim Strmsg As String
        Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
        DelReport_All_TempRst2
        Temp_Table2.Delete
        ComTest_Name1 = ""
        txtTest_Result1 = ""
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_TempTable3_Click()
    If ComTest_Name2 = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
    If Temp_Table3.RecordCount <= 0 Then Exit Sub
    
    If Trim(ComTest_Name2.Text) = "" Then
        MsgBox "You didn't select the the Test Name"
        DataGrid4.SetFocus
        Exit Sub
    Else
        Dim Strmsg As String
        Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
        DelReport_All_TempRst3
        Temp_Table3.Delete
        ComTest_Name2 = ""
        txtTest_Result2 = ""
        
        End If
        
    End If
End Sub

Private Sub cmdPreview_Click()
    CRViewer1_MODE = 21
    RptViewer.Show
End Sub

Private Sub cmdPrint_Click()
    '==========direct print==========================
            
            Dim Report21 As New Urine
            Dim StrPat_ID As String
            Dim strM_Code As String
            Dim strS_Code As String
            
            StrPat_ID = rUrine1.txtPat_ID
            strM_Code = rUrine1.txtM_Code
            strS_Code = rUrine1.txtS_Code
            
            Report21.DiscardSavedData
            RS.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report21.Database.SetDataSource RS
            
            Report21.PrintOut
            RS.Close

    '====================================
End Sub

Private Sub cmdSave_Click()
'-----validation check---------------------
    If Trim(txtPat_ID) = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtS_Code)) = 0 Then
        MsgBox "Test Code mandatory"
        txtS_Code.SetFocus
        Exit Sub
    End If
'-----end validation check--------------------------------------------------

''\\\\\\\\\\INSERT and UPDATE from Temp_rst1\\\\\\\\\\\\\
    Adodc2.connectionstring = strcn.Connection
'    Adodc2.RecordSource = "Report_All_SELECT3 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc2.RecordSource = "Report_All_SELECT 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "3" + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        Del_All_Report_All_TempRst1
        InsReport_All_TempRst1
'        InsReport_All_TempRst2
'        InsReport_All_TempRst3
        'MsgBox "Updated1"
    Else
        InsReport_All_TempRst1
'        InsReport_All_TempRst2
'        InsReport_All_TempRst3
        'MsgBox "Inserted1"
    End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
''\\\\\\\\\\INSERT and UPDATE from Temp_rst2\\\\\\\\\\\\\
    Adodc2.connectionstring = strcn.Connection
'    Adodc2.RecordSource = "Report_All_SELECT3 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc2.RecordSource = "Report_All_SELECT 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "4" + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        Del_All_Report_All_TempRst1
'        InsReport_All_TempRst1
        InsReport_All_TempRst2
'        InsReport_All_TempRst3
        'MsgBox "Updated2"
    Else
'        InsReport_All_TempRst1
        InsReport_All_TempRst2
'        InsReport_All_TempRst3
        'MsgBox "Inserted2"
    End If
''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
''\\\\\\\\\\INSERT and UPDATE from Temp_rst2\\\\\\\\\\\\\
    Adodc2.connectionstring = strcn.Connection
'    Adodc2.RecordSource = "Report_All_SELECT3 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc2.RecordSource = "Report_All_SELECT 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "5" + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        Del_All_Report_All_TempRst1
'        InsReport_All_TempRst1
'        InsReport_All_TempRst2
        InsReport_All_TempRst3
        'MsgBox "Updated3"
    Else
'        InsReport_All_TempRst1
'        InsReport_All_TempRst2
        InsReport_All_TempRst3
        'MsgBox "Inserted3"
    End If
''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Temp_rst1 (False)
    Temp_rst2 (False)
    Temp_rst3 (False)


    
End Sub
Private Sub cmdShow_Click()
If cmdSave.Enabled = False Then Exit Sub

        If Len(txtPat_ID.Text) = 0 Then
        MsgBox "Patient ID mandatory"
        txtPat_ID.SetFocus
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
        txtPat_ID.SetFocus
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
    ComTest_Name.Text = DataGrid2.Columns(0)
    txtTest_Result.Text = DataGrid2.Columns(1)
    
End Sub
Private Sub DataGrid3_DblClick()
    ComTest_Name1.Text = DataGrid3.Columns(0)
    txtTest_Result1.Text = DataGrid3.Columns(1)
End Sub
Private Sub DataGrid4_DblClick()
    ComTest_Name2.Text = DataGrid4.Columns(0)
    txtTest_Result2.Text = DataGrid4.Columns(1)
End Sub

Private Sub Form_Click()
    DataGrid1.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    If rUrine1.DataGrid1.Visible = True Then
        rUrine1.DataGrid1.Visible = False
    Else
        Unload Me
    End If
    End If
End Sub

Private Sub Form_Load()
    
    Adodc8.connectionstring = strcn.Connection
    Adodc8.RecordSource = "exec m_name_select 2,'" + "URINE" + "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
        txtM_Code = Adodc8.Recordset!m_code
    Else
        MsgBox "Inserted incurrect head name, first you have to insert currect name from TEST INFORMATION form then open this screen again"
        txtPat_ID.Enabled = False
        cmdSave.Enabled = False
    End If


    Dt.value = Now
    Temp_rst1 (True)
    Temp_rst2 (True)
    Temp_rst3 (True)
    
'------select data from test_result------
     GetTestName
     GetTestName1
     GetTestName2
'------end------------------------------
    
DataGrid2.Columns(0).Width = 3415
DataGrid2.Columns(1).Width = 5020
DataGrid3.Columns(0).Width = 3415
DataGrid3.Columns(1).Width = 5020
DataGrid4.Columns(0).Width = 3415
DataGrid4.Columns(1).Width = 5020
    
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

If Len(Trim(txtPat_ID.Text)) = 0 Then Exit Sub
    
Adodc5.connectionstring = strcn.Connection
Adodc5.RecordSource = "exec Pro_FLUSH 6,'" + txtPat_ID + "'"
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
                    MsgBox "Invalied PATIENT ID"
                    txtPat_ID = ""
                    txtPat_ID.SetFocus
                    
                End If
        '===============================================
End If


End Sub
Private Sub DataGrid1_DblClick()
    'txtM_Code.Text = DataGrid1.Columns(0)
    txtS_Code.Text = DataGrid1.Columns(1)
    txtS_Name.Text = DataGrid1.Columns(2)
    txtSpecimen.SetFocus
    DataGrid1.Visible = False
End Sub
Private Sub txtRef_Range_LostFocus()

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


If Len(txtS_Code.Text) = 0 Then Exit Sub
If cmdSave.Enabled = False Then Exit Sub

If cmdSave.Enabled = False Then Exit Sub

If Trim(txtPat_ID) = "" Then
    MsgBox "Patient ID mandatory"
    txtPat_ID.SetFocus
    Exit Sub
End If

'If Len(Trim(txtS_Code)) = 0 Then Exit Sub
         
    Adodc6.connectionstring = strcn.Connection
    Adodc6.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "3" + "'"
    Adodc6.Refresh
    
    If Adodc6.Recordset.RecordCount > 0 Then
        DataGrid1.Visible = False
    
         txtSpecimen = Adodc6.Recordset!Field1
         txtSpecimen_dt_Time = Adodc6.Recordset!Field2
         txtN_Exam = Adodc6.Recordset!Field3
         txtUsed_tech = Adodc6.Recordset!Field4
         Dt.value = Adodc6.Recordset!Dt

'++++++++++for show feild5 to txtNote FROM TYPE 3 +++++++++++
    Adodc8.connectionstring = strcn.Connection
    Adodc8.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "3" + "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
    txtNote = Adodc8.Recordset!Field15
    txtPh_Exam = Adodc8.Recordset!Field5
    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++for show feild5 to CHAMICAL EXAMINATION FROM TYPE 4 +++++++++++
    Adodc10.connectionstring = strcn.Connection
    Adodc10.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "4" + "'"
    Adodc10.Refresh
    If Adodc10.Recordset.RecordCount > 0 Then
    txtChm_Exam = Adodc10.Recordset!Field5
    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++for show feild5 to MICROSCOPIC EXAMINATION FROM TYPE 5 +++++++++++
    Adodc11.connectionstring = strcn.Connection
    Adodc11.RecordSource = "exec Report_All_Select1 1,'" & Trim(txtPat_ID.Text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + "5" + "'"
    Adodc11.Refresh
    If Adodc11.Recordset.RecordCount > 0 Then
    txtMicro_Exam = Adodc11.Recordset!Field5
    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
         
'/////////show in Temp_rst1//////////////
        con.connectionstring = strcn.Connection
        con.Open
        Temp_Table_Helper1.Open "select * from report_all where pat_id='" + txtPat_ID + "' and s_code='" + txtS_Code + "'and m_code='" + txtM_Code + "'and type='" + "3" + "'", con
        
          While Temp_Table_Helper1.EOF = False
                Temp_Table1.AddNew
                Temp_Table1!Test_Name = Temp_Table_Helper1!Field6
                Temp_Table1!test_result = Temp_Table_Helper1!Field7
                Temp_Table_Helper1.MoveNext
            Wend
        DataGrid2.Refresh
        Temp_Table_Helper1.Close
        con.Close
'/////////end show in Temp_rst1////////////////////////////
'/////////show in Temp_rst2//////////////
        con.connectionstring = strcn.Connection
        con.Open
        Temp_Table_Helper2.Open "select * from report_all where pat_id='" + txtPat_ID + "' and s_code='" + txtS_Code + "'and m_code='" + txtM_Code + "'and type='" + "4" + "'", con
        
          While Temp_Table_Helper2.EOF = False
                Temp_Table2.AddNew
                Temp_Table2!Test_Name1 = Temp_Table_Helper2!Field6
                Temp_Table2!test_result1 = Temp_Table_Helper2!Field7
                Temp_Table_Helper2.MoveNext
            Wend
        DataGrid3.Refresh
        Temp_Table_Helper2.Close
        con.Close
'/////////end show in Temp_rst2////////////////////////////
'/////////show in Temp_rst3//////////////
        con.connectionstring = strcn.Connection
        con.Open
        Temp_Table_Helper3.Open "select * from report_all where pat_id='" + txtPat_ID + "' and s_code='" + txtS_Code + "'and m_code='" + txtM_Code + "'and type='" + "5" + "'", con
        
          While Temp_Table_Helper3.EOF = False
                Temp_Table3.AddNew
                Temp_Table3!Test_Name2 = Temp_Table_Helper3!Field6
                Temp_Table3!test_result2 = Temp_Table_Helper3!Field7
                Temp_Table_Helper3.MoveNext
            Wend
        DataGrid4.Refresh
        Temp_Table_Helper3.Close
        con.Close
'/////////end show in Temp_rst3////////////////////////////

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
    
DataGrid2.Columns(0).Width = 3415
DataGrid2.Columns(1).Width = 5020
DataGrid3.Columns(0).Width = 3415
DataGrid3.Columns(1).Width = 5020
DataGrid4.Columns(0).Width = 3415
DataGrid4.Columns(1).Width = 5020

       
End Sub
Public Sub Temp_rst1(temp_open1 As Boolean)

    If temp_open1 = False Then
        Temp_Table1.Close
        temp_open1 = True
    End If
    
    If temp_open1 = True Then
        With Temp_Table1
            .Fields.Append "Test_Name", adVarChar, 500
            .Fields.Append "Test_Result", adVarChar, 500
            .LockType = adLockOptimistic
            .Open
            temp_open1 = False
        End With
            Set DataGrid2.DataSource = Temp_Table1
            DataGrid2.ReBind
            DataGrid2.Refresh
            
    End If
End Sub
Public Sub Temp_rst2(temp_open2 As Boolean)

    If temp_open2 = False Then
        Temp_Table2.Close
        temp_open2 = True
    End If
    
    If temp_open2 = True Then
        With Temp_Table2
            .Fields.Append "Test_Name1", adVarChar, 500
            .Fields.Append "Test_Result1", adVarChar, 500
            .LockType = adLockOptimistic
            .Open
            temp_open2 = False
        End With
            Set DataGrid3.DataSource = Temp_Table2
            DataGrid3.ReBind
            DataGrid3.Refresh
            
    End If
End Sub
Public Sub Temp_rst3(temp_open3 As Boolean)

    If temp_open3 = False Then
        Temp_Table3.Close
        temp_open3 = True
    End If
    
    If temp_open3 = True Then
        With Temp_Table3
            .Fields.Append "Test_Name2", adVarChar, 500
            .Fields.Append "Test_Result2", adVarChar, 500
            .LockType = adLockOptimistic
            .Open
            temp_open3 = False
        End With
            Set DataGrid4.DataSource = Temp_Table3
            DataGrid4.ReBind
            DataGrid4.Refresh
            
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
            "','" + txtSpecimen + _
            "','" + txtSpecimen_dt_Time + _
            "','" + txtN_Exam + _
            "','" + txtUsed_tech + _
            "','" + txtPh_Exam + _
            "','" + Temp_Table1!Test_Name + _
            "','" + Temp_Table1!test_result + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + txtNote + _
            "','" + u_id + _
            "','" + Format(Dt, "yyyy-mm-dd") + _
            "','" + "3" + "'"
            cmd.Execute
            Temp_Table1.MoveNext
    Wend
    con.Close
End Sub
Private Sub InsReport_All_TempRst2()
    
    Temp_Table2.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    While Temp_Table2.EOF = False
    
          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
            "','" + txtM_Code + _
            "','" + txtS_Code + _
            "','" + txtSpecimen + _
            "','" + txtSpecimen_dt_Time + _
            "','" + txtN_Exam + _
            "','" + txtUsed_tech + _
            "','" + txtChm_Exam + _
            "','" + Temp_Table2!Test_Name1 + _
            "','" + Temp_Table2!test_result1 + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + txtNote + _
            "','" + u_id + _
            "','" + Format(Dt, "yyyy-mm-dd") + _
            "','" + "4" + "'"
            cmd.Execute
            Temp_Table2.MoveNext
    Wend
    con.Close
End Sub
Private Sub InsReport_All_TempRst3()
    Temp_Table3.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    While Temp_Table3.EOF = False
    
          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
            "','" + txtM_Code + _
            "','" + txtS_Code + _
            "','" + txtSpecimen + _
            "','" + txtSpecimen_dt_Time + _
            "','" + txtN_Exam + _
            "','" + txtUsed_tech + _
            "','" + txtMicro_Exam + _
            "','" + Temp_Table3!Test_Name2 + _
            "','" + Temp_Table3!test_result2 + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + txtNote + _
            "','" + u_id + _
            "','" + Format(Dt, "yyyy-mm-dd") + _
            "','" + "5" + "'"
            cmd.Execute
            Temp_Table3.MoveNext
    Wend
    con.Close
End Sub
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
'            "','" + txtUsed_tech + _
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
Private Sub DelReport_All_TempRst1()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.Text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "3" + "'"
            cmd.Execute
    con.Close
End Sub
Private Sub DelReport_All_TempRst2()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.Text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "4" + "'"
            cmd.Execute
    con.Close
End Sub
Private Sub DelReport_All_TempRst3()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.Text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "5" + "'"
            cmd.Execute
    con.Close
End Sub


Private Sub Del_All_Report_All_TempRst1()
   
    Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    While Temp_Table1.EOF = False
            cmd.CommandText = "exec Report_All_Delete2 1,'" + Trim(txtPat_ID.Text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
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
Private Sub txtTest_Result_LostFocus()
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
        Temp_Table1!test_result = txtTest_Result
        DataGrid2.Refresh
'+++++++++++++++++++++++++++++++++++++++
'    DataGrid2.Columns(0).Width = 2000
'    DataGrid2.Columns(1).Width = 1000
    
    ComTest_Name.SetFocus
    
ComTest_Name = ""
txtTest_Result = ""


DataGrid2.Columns(0).Width = 3415
DataGrid2.Columns(1).Width = 5020

End Sub
Private Sub txtTest_Result1_GotFocus()
    GetResult1
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
    GetResult2
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
'    DataGrid4.Columns(0).Width = 2000
'    DataGrid4.Columns(1).Width = 1000
    
    ComTest_Name2.SetFocus
    
ComTest_Name2 = ""
txtTest_Result2 = ""


DataGrid4.Columns(0).Width = 3415
DataGrid4.Columns(1).Width = 5020

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
    txtSpecimen = ""
    txtUsed_tech = ""
    txtN_Exam = ""
    txtNote = ""
    Dt.value = Now
    
End Sub
Private Sub GetTestName()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select1 1,'" + "08A" + "'"
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
  Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Name + "','" + "08A" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    txtTest_Result = Adodc7.Recordset!test_result
    txtUnit = Adodc7.Recordset!unit
   ' txtRef_Range = Adodc7.Recordset!ref_range
    End If
End Sub
Private Sub GetTestName1()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select1 1,'" + "08B" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
       Do Until Adodc7.Recordset.EOF
          ComTest_Name1.AddItem Adodc7.Recordset!Test_Name
       Adodc7.Recordset.MoveNext
       Loop
    End If
End Sub
Private Sub GetResult1()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Name1 + "','" + "08B" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    txtTest_Result1 = Adodc7.Recordset!test_result
    txtUnit1 = Adodc7.Recordset!unit
   ' txtRef_Range = Adodc7.Recordset!ref_range
    End If
End Sub
Private Sub GetTestName2()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select1 1,'" + "08C" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
       Do Until Adodc7.Recordset.EOF
          ComTest_Name2.AddItem Adodc7.Recordset!Test_Name
       Adodc7.Recordset.MoveNext
       Loop
    End If
End Sub
Private Sub GetResult2()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Name2 + "','" + "08C" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    txtTest_Result2 = Adodc7.Recordset!test_result
    'txtUnit2 = Adodc7.Recordset!unit
   ' txtRef_Range = Adodc7.Recordset!ref_range
    End If
End Sub


