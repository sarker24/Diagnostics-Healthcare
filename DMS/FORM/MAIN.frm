VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMAIN 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic management system [Unique Diagnostic Center]"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15270
   Icon            =   "MAIN.frx":0000
   Picture         =   "MAIN.frx":0ECA
   ScaleHeight     =   10860
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   150
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":22F2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":22F89
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":22FE7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   11160
      Top             =   1800
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   10530
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2206
            MinWidth        =   2206
            Text            =   "User Name"
            TextSave        =   "User Name"
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "User Name"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Object.ToolTipText     =   "Log in Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2188
            MinWidth        =   2188
            Object.ToolTipText     =   "Log in Time"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   17727
            Text            =   "Software Developed by ""MAS IT SOLUTIONS"". Hot Line: +880-2-8056691, 01915682291, 01714589268"
            TextSave        =   "Software Developed by ""MAS IT SOLUTIONS"". Hot Line: +880-2-8056691, 01915682291, 01714589268"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   3
      Left            =   240
      MouseIcon       =   "MAIN.frx":23045
      MousePointer    =   99  'Custom
      ToolTipText     =   "Open Doctor Inforamtion form"
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Doctor_Commision 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Payment Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      MouseIcon       =   "MAIN.frx":2334F
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "press CTRL + D"
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   4
      Left            =   240
      MouseIcon       =   "MAIN.frx":23659
      MousePointer    =   99  'Custom
      ToolTipText     =   "Open Doctor Inforamtion form"
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label lblCommission 
      BackStyle       =   0  'Transparent
      Caption         =   "Reffered Commission Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      MouseIcon       =   "MAIN.frx":23963
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "press CTRL + D"
      Top             =   6120
      Width           =   3855
   End
   Begin VB.Label Doctor_Profile 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      MouseIcon       =   "MAIN.frx":23C6D
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "press CTRL + D"
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Label Test_Information 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Informations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      MouseIcon       =   "MAIN.frx":23F77
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "press CTRL + T"
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Investigation 
      BackStyle       =   0  'Transparent
      Caption         =   "Investigation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      MouseIcon       =   "MAIN.frx":24281
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "press CTRL + I"
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   240
      MouseIcon       =   "MAIN.frx":2458B
      MousePointer    =   99  'Custom
      ToolTipText     =   "Open Investigation form"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   240
      MouseIcon       =   "MAIN.frx":24895
      MousePointer    =   99  'Custom
      ToolTipText     =   "Open Test List  form"
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   240
      MouseIcon       =   "MAIN.frx":24B9F
      MousePointer    =   99  'Custom
      ToolTipText     =   "Open Doctor Inforamtion form"
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick A Task"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Menu mnuseparate 
      Caption         =   "-----"
   End
   Begin VB.Menu mnuENTRY 
      Caption         =   "&Setup"
      Begin VB.Menu mnuCompany_Information 
         Caption         =   "&Company Information"
      End
      Begin VB.Menu mnuEmp_Info 
         Caption         =   "Employee Information"
      End
      Begin VB.Menu mnuSup_Info 
         Caption         =   "Supplier Information"
      End
      Begin VB.Menu mnuItem_Info 
         Caption         =   "Item Information"
      End
      Begin VB.Menu mnuStock_In 
         Caption         =   "Stock In"
      End
      Begin VB.Menu mnuItem_Issue 
         Caption         =   "Item Issue"
      End
      Begin VB.Menu mnuLeave_Setup 
         Caption         =   "Leave Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu frmEmp_Leave 
         Caption         =   "Employee Leave"
         Visible         =   0   'False
      End
      Begin VB.Menu mmdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCREATE_USER 
         Caption         =   "Create User"
      End
      Begin VB.Menu mnuChange_Password 
         Caption         =   "Change Password"
      End
      Begin VB.Menu MNUPAtient_Info_VAT 
         Caption         =   "Patient Information New"
      End
      Begin VB.Menu mnuDoctor_Information 
         Caption         =   "&Doctor's Information"
      End
      Begin VB.Menu MNUComm_percent 
         Caption         =   "Commission percent"
      End
      Begin VB.Menu mnuTest_Information 
         Caption         =   "&Test Information"
      End
      Begin VB.Menu mnuTestResult 
         Caption         =   "Test Result"
      End
      Begin VB.Menu mm 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAuthority 
         Caption         =   "Authority"
      End
      Begin VB.Menu mnuD_Backup 
         Caption         =   "Data Backup"
      End
      Begin VB.Menu mnuPay_Edit 
         Caption         =   "Payment Modification"
      End
      Begin VB.Menu mnuTest_Monidy 
         Caption         =   "Test Modification"
      End
      Begin VB.Menu mnuDisc_Edit 
         Caption         =   "Discount Modification"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font Setup"
      End
      Begin VB.Menu mnuVAT_Setup 
         Caption         =   "VAT Setup"
      End
      Begin VB.Menu MNUDOT 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEXIT 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "Transaction"
      Begin VB.Menu mnuPatient_Information 
         Caption         =   "&Patient Information"
      End
      Begin VB.Menu mnuDPEntry 
         Caption         =   "Doctor's Patient Entry"
      End
      Begin VB.Menu mnuDue_Coll 
         Caption         =   "Dues Collection"
      End
      Begin VB.Menu mnuDPayment 
         Caption         =   "Doctor Payment"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuBOOTH_STATUS 
         Caption         =   "Booth Status"
      End
      Begin VB.Menu mnu_Test_Information 
         Caption         =   "Test Information"
      End
      Begin VB.Menu mnuStock_Status 
         Caption         =   "Stock Status"
      End
      Begin VB.Menu mnuDoctor_Info 
         Caption         =   "Doctor's Information"
      End
      Begin VB.Menu mnuDVStatement 
         Caption         =   "Doctor Visit Statement"
      End
      Begin VB.Menu mnuRDPDetails 
         Caption         =   "Refferred Doctor's Payment Details"
      End
      Begin VB.Menu mnuCWIStatement 
         Caption         =   "Consultant wise Income Statement"
      End
      Begin VB.Menu mnuDPStatement 
         Caption         =   "Doctor Payment Statement"
      End
      Begin VB.Menu mnuDoc_Due_pat 
         Caption         =   "Doctor's Due Patient"
      End
      Begin VB.Menu mmuSalesManPerformance 
         Caption         =   "Marketing Executive Performance"
      End
      Begin VB.Menu mnuNEW_DOC_INFO 
         Caption         =   "New Doctor's Information"
      End
      Begin VB.Menu mnuAReport 
         Caption         =   "Accounts Report"
         Begin VB.Menu mnuCBook 
            Caption         =   "Income Statement"
         End
         Begin VB.Menu mnuBBook 
            Caption         =   "Bank Book"
         End
         Begin VB.Menu mnuGLedger 
            Caption         =   "General Ledger"
         End
         Begin VB.Menu mnuTBalance 
            Caption         =   "Trial Balance"
         End
         Begin VB.Menu mnuPLAccounts 
            Caption         =   "Profit & Lose Accounts"
         End
         Begin VB.Menu mnuBSheet 
            Caption         =   "Balance Sheet"
         End
      End
      Begin VB.Menu mnuDD 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDaily_Statement 
         Caption         =   "Daily Statement"
      End
      Begin VB.Menu mnuGroup_Test 
         Caption         =   "Groupwise Test Statement"
      End
      Begin VB.Menu mnuTWStatemetn 
         Caption         =   "Department and Test wise Statement"
      End
      Begin VB.Menu mnuAdv_Coll 
         Caption         =   "Collection Information"
      End
      Begin VB.Menu mnuPat_Type 
         Caption         =   "Patient Type"
      End
      Begin VB.Menu mnuPat_info 
         Caption         =   "Patient Information"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTest_Reports 
      Caption         =   "&Lab Reports"
      Begin VB.Menu mnuHaem 
         Caption         =   "Haematology Analysis"
      End
      Begin VB.Menu mnuBio 
         Caption         =   "Bio Chemical"
      End
      Begin VB.Menu mnuImmu 
         Caption         =   "Immunology"
      End
      Begin VB.Menu mnuHepa 
         Caption         =   "Hepatitis Profile"
      End
      Begin VB.Menu mnuHormo 
         Caption         =   "Hormone"
      End
      Begin VB.Menu mnuTM 
         Caption         =   "Tumour Marker"
      End
      Begin VB.Menu mnuMicroB 
         Caption         =   "Microbiology"
      End
      Begin VB.Menu mnuMicro_Spec 
         Caption         =   "Microbiology Special"
      End
      Begin VB.Menu mnuUrine 
         Caption         =   "Urine"
      End
      Begin VB.Menu mnuUrine_Spec 
         Caption         =   "Urine Special"
      End
      Begin VB.Menu mnuStool 
         Caption         =   "Stool"
      End
      Begin VB.Menu mnuStool_Spec 
         Caption         =   "Stool Special"
      End
      Begin VB.Menu mnuBodyF 
         Caption         =   "Body Fluid"
      End
      Begin VB.Menu mnuX_Ray 
         Caption         =   "X-Ray"
      End
      Begin VB.Menu mnuEndoscopy 
         Caption         =   "Endoscopy"
      End
      Begin VB.Menu mnuEchoCar 
         Caption         =   "Echocardiagraphy"
      End
      Begin VB.Menu mnuCTscan 
         Caption         =   "C.T. Scan"
      End
      Begin VB.Menu mnuMamo 
         Caption         =   "Mammography"
      End
      Begin VB.Menu mnuDrug 
         Caption         =   "Drug"
      End
      Begin VB.Menu mnuUltra 
         Caption         =   "Ultrasonogram"
      End
      Begin VB.Menu mnuHistoP 
         Caption         =   "Histopathology"
      End
      Begin VB.Menu mnuPaps 
         Caption         =   "Paps"
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "&Accounts"
      Begin VB.Menu mnuHAccounts 
         Caption         =   "Head of Accounts"
      End
      Begin VB.Menu mnuVEntry 
         Caption         =   "Voucher Entry"
      End
      Begin VB.Menu mnuJournalEntry 
         Caption         =   "Journal Entry"
         WindowList      =   -1  'True
      End
      Begin VB.Menu mnuBankEntry 
         Caption         =   "Bank Entry"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "&Calculator"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuLPrinting 
         Caption         =   "Level Printing"
      End
      Begin VB.Menu mnuEnvelope 
         Caption         =   "En&velope"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About "
      End
   End
   Begin VB.Menu mExit 
      Caption         =   "&Log off "
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Doctor_Commision_Click()
strScr_Name = "rDoc_Pay"
        Authority
    If strAllow = "YES" Then
        rDoc_Pay.Show vbModal
    End If
End Sub

Private Sub Doctor_Profile_Click()
strScr_Name = "frmDoctor_Info"
        Authority
    If strAllow = "YES" Then
    frmDoctor_Info.Show vbModal
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Me.StatusBar1.Panels(2) = u_id
Me.StatusBar1.Panels(3) = Date 'Format "dd-MM-yyyy"
Me.StatusBar1.Panels(4) = Time
'    With ShockwaveFlash1
'         .Movie = App.path + "\prime.swf"
'         .Loop = False
'         .Play
'    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'ques = MsgBox("Do you want to exit the Application", vbQuestion + vbYesNo, "Restaurant Management System.....")
'If ques = vbYes Then
'    End
'Else
'    Cancel = 1
'End If
End Sub

Private Sub frmEmp_Leave_Click()
    strScr_Name = "frmCompany_Info"
    Authority
    If strAllow = "YES" Then
    frmLeave.Show vbModal
    End If
    
End Sub

Private Sub Investigation_Click()
strScr_Name = "frmPatient_Info"
        Authority
    If strAllow = "YES" Then
    frmPatient_Info.Show vbModal
    End If
End Sub

Private Sub lblCommission_Click()
strScr_Name = "frmCommission_Per"
        Authority
    If strAllow = "YES" Then
    frmCommission_Per.Show vbModal
    End If
End Sub

Private Sub mExit_Click()
    Dim res As VbMsgBoxResult
    res = MsgBox("Are you sure you want to log off?", vbYesNo + vbQuestion)
    If res = vbYes Then
    Unload Me
    'frmLogin.Show
    frmLogIn.Show
    frmLogIn.Txtuserid = ""
    frmLogIn.Txtpass = ""
    Else
    End If
End Sub

Private Sub mmuSalesManPerformance_Click()
'Dim F As New rSalesManPerformance
'strScr_Name = "rptMExecutive"
strScr_Name = "rPat_Info"
        Authority
    If strAllow = "YES" Then
        rPat_Info.Show vbModal
    End If
'F.Show 1
End Sub

Private Sub mnu_D_P_Summery_Click()
strScr_Name = "rDoc_Pay_Summery"
        Authority
    If strAllow = "YES" Then
        rDoc_Pay_Summery.Show vbModal
    End If
End Sub

Private Sub mnu_Test_Information_Click()
    strScr_Name = "RptTest_Info"
        Authority
    If strAllow = "YES" Then
       RptTest_Info.Show vbModal
    End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuAdv_Coll_Click()
    strScr_Name = "rAdv_Coll"
        Authority
    If strAllow = "YES" Then
        rAdv_Coll.Show vbModal
    End If
End Sub

Private Sub mnuAuthority_Click()
    strScr_Name = "frmUser_Authority"
        Authority
    If strAllow = "YES" Then
    frmUser_Authority.Show vbModal
    End If
End Sub

Private Sub mnuBankEntry_Click()
strScr_Name = "frmRefDoc"
    Authority
    If strAllow = "YES" Then
    frmRefDoc.Show vbModal
    End If

End Sub

Private Sub mnuBBook_Click()
strScr_Name = "RptBankBook"
        Authority
    If strAllow = "YES" Then
    RptBankBook.Show vbModal
    End If

End Sub

Private Sub mnuBio_Click()
    'strScr_Name = "rBio_Chamical"
    'Authority
    'If strAllow = "YES" Then
    rBio_Chamical.Show vbModal
    'End If
End Sub

Private Sub mnuBodyF_Click()
    'strScr_Name = "rBody_Fluid"
    'Authority
    'If strAllow = "YES" Then
    rBody_Fluid.Show vbModal
    'End If
End Sub

Private Sub mnuBOOTH_STATUS_Click()
    strScr_Name = "rBooth_User_Info"
        Authority
    If strAllow = "YES" Then
    rBooth_User_Info.Show vbModal
    End If

End Sub

Private Sub mnuBSheet_Click()
'strScr_Name = "frmBSheet"
'    Authority
'    If strAllow = "YES" Then
'    frmBSheet.Show vbModal
'    End If

End Sub

Private Sub mnuCalculator_Click()
On Error Resume Next
   Shell "calc.exe"
End Sub

Private Sub mnuCBook_Click()
strScr_Name = "RptCashBook"
        Authority
    If strAllow = "YES" Then
    RptCashBook.Show vbModal
    End If

End Sub

Private Sub mnuChange_Password_Click()
    strScr_Name = "frmChange_Password"
        Authority
    If strAllow = "YES" Then
    frmChange_Password.Show vbModal
    End If

End Sub

Private Sub mnuComm_Pay_Edit_Click()
    strScr_Name = "frmCommPayEdit"
    Authority
    If strAllow = "YES" Then
    frmCommPayEdit.Show vbModal
    End If
End Sub

Private Sub MNUComm_percent_Click()
    strScr_Name = "frmCommission_Per"
        Authority
    If strAllow = "YES" Then
    frmCommission_Per.Show vbModal
    End If

End Sub

Private Sub mnuCOMMISSION_EDIT_Click()
    strScr_Name = "frmCommEdit"
    Authority
    If strAllow = "YES" Then
    frmCommEdit.Show vbModal
    End If
    
End Sub

Private Sub mnuCompany_Information_Click()

    strScr_Name = "frmCompany_Info"
        Authority
    If strAllow = "YES" Then
    frmCompany_Info.Show vbModal
    End If
    
End Sub

Private Sub mnuCREATE_USER_Click()
   strScr_Name = "frmCreate_User"
   Authority
   If strAllow = "YES" Then
       frmCreate_User.Show vbModal
   End If
End Sub

Private Sub mnuCTscan_Click()
   'strScr_Name = "rCT_SCAN"
   'Authority
   'If strAllow = "YES" Then
    rCT_SCAN.Show vbModal
   'End If
End Sub

Private Sub mnuCWIStatement_Click()
strScr_Name = "RptConsultant"
        Authority
   If strAllow = "YES" Then
        RptConsultant.Show vbModal
   End If
End Sub

Private Sub mnuD_Backup_Click()
frmBackUp.Show vbModal
'    'ProgressBar1.Visible = True
'    Kill "E:\Prime_Back\*.txt"
'
'    Shell ("E:\Prime_Back\prime_out.bat")
'    ', vbHide
'    'ProgressBar1.Visible = False
    
End Sub

Private Sub mnuDaily_Statement_Click()
   strScr_Name = "rDaily_Statement"
        Authority
   If strAllow = "YES" Then
        rDaily_Statement.Show vbModal
   End If
End Sub

Private Sub mnuDData_Click()
strScr_Name = "frmDataDelete"
        Authority
    If strAllow = "YES" Then
        frmDataDelete.Show vbModal
    End If
End Sub

Private Sub mnuDisc_Edit_Click()
    strScr_Name = "frmDisc_Edit"
        Authority
    If strAllow = "YES" Then
    frmDisc_Edit.Show vbModal
    End If
End Sub

Private Sub mnuDOC_COMM_DETAILS_Click()
    strScr_Name = "rDoc_Pay"
        Authority
    If strAllow = "YES" Then
        rDoc_Pay.Show vbModal
    End If

End Sub
Private Sub mnuDoc_Due_pat_Click()
    strScr_Name = "rDoc_Due_Pat"
    Authority
    If strAllow = "YES" Then
    rDoc_Due_Pat.Show vbModal
    End If
End Sub
Private Sub mnuDoctor_Info_Click()
    strScr_Name = "RptDoctor_Info"
        Authority
    If strAllow = "YES" Then
    RptDoctor_Info.Show vbModal
    End If
End Sub
Private Sub mnuDoctor_Information_Click()
    strScr_Name = "frmDoctor_Info"
        Authority
    If strAllow = "YES" Then
    frmDoctor_Info.Show vbModal
    End If

End Sub

Private Sub mnuDPayment_Click()
'strScr_Name = "frmCommissionPay"
'    Authority
'
'    If strAllow = "YES" Then
    frmCommissionPay.Show vbModal
'    End If
End Sub

Private Sub mnuDPEntry_Click()
strScr_Name = "frmDoctorSerial"
    Authority
    
    If strAllow = "YES" Then
    frmDoctorSerial.Show vbModal
    End If
End Sub

Private Sub mnuDPStatement_Click()
'strScr_Name = "rDrug"
   'Authority
   'If strAllow = "YES" Then
    RptDPayment.Show vbModal
    'End If
End Sub

Private Sub mnuDrug_Click()
   'strScr_Name = "rDrug"
   'Authority
   'If strAllow = "YES" Then
    rDrug.Show vbModal
    'End If
End Sub

Private Sub mnuDue_Coll_Click()
    strScr_Name = "frmPat_Info_Due"
    Authority
    If strAllow = "YES" Then
        frmPat_Info_Due.Show vbModal
    End If
    
End Sub

Private Sub mnuDVStatement_Click()
strScr_Name = "RptDVStatement"
    Authority
    If strAllow = "YES" Then
    RptDVStatement.Show vbModal
    End If
    
End Sub

Private Sub mnuEchoCar_Click()
   'strScr_Name = "rEchocardiography"
   'Authority
   'If strAllow = "YES" Then
   rEchocardiography.Show vbModal
   'End If
End Sub

Private Sub mnuEmp_Info_Click()
    strScr_Name = "frmCompany_Info"
    Authority
    If strAllow = "YES" Then
    frmEmp_Info.Show vbModal
    End If
    
End Sub

Private Sub mnuEndoscopy_Click()
    'strScr_Name = "rEndoscopy"
    'Authority
    'If strAllow = "YES" Then
    rEndoscopy.Show vbModal
    'End If
End Sub

Private Sub mnuEnvelope_Click()
    rEnvelope.Show vbModal
End Sub

Private Sub mnuEXIT_Click()
    End
End Sub

Private Sub mnuFont_Click()
    'strScr_Name = "frmFont"
    'Authority
    'If strAllow = "YES" Then
    frmFont.Show vbModal
    'End If
End Sub



Private Sub mnuGroup_Test_Click()
    'strScr_Name = "rDaily_Test"
    'Authority
    
    'If strAllow = "YES" Then
    rDaily_Test.Show vbModal
    'End If

End Sub

Private Sub mnuHAccounts_Click()
strScr_Name = "frmAccountsHead"
    Authority
    
    If strAllow = "YES" Then
    frmAccountsHead.Show vbModal
    End If
End Sub

Private Sub mnuHaem_Click()
    'strScr_Name = "rHaematology"
    'Authority
    
    'If strAllow = "YES" Then
    rHaematology.Show vbModal
    'End If
End Sub

Private Sub mnuHepa_Click()
    'strScr_Name = "rHepatitis"
    'Authority
    'If strAllow = "YES" Then
    rHepatitis.Show vbModal
    'End If
End Sub

Private Sub mnuHistoP_Click()
    'strScr_Name = "rHepatitis"
    'Authority
    'If strAllow = "YES" Then
    rHistopathology.Show vbModal
    'End If
End Sub
Private Sub mnuHormo_Click()
    'strScr_Name = "rHormone"
    'Authority
    'If strAllow = "YES" Then
        rHormone.Show vbModal
    'End If
End Sub

Private Sub mnuImmu_Click()
    'strScr_Name = "rImmunology"
    'Authority
    'If strAllow = "YES" Then
    rImmunology.Show vbModal
    'End If
End Sub

Private Sub mnuItem_Info_Click()
    strScr_Name = "frmCompany_Info"
    Authority
    If strAllow = "YES" Then
    frmItem_Info.Show vbModal
    End If
End Sub
Private Sub mnuItem_Issue_Click()
    strScr_Name = "frmCompany_Info"
    Authority
    If strAllow = "YES" Then
    frmStock_Out.Show vbModal
    End If
End Sub
Private Sub mnuLeave_as_cash_Click()
    strScr_Name = "frmLeave_as_Cash"
    Authority
    If strAllow = "YES" Then
    frmLeave_as_Cash.Show vbModal
    End If
End Sub
Private Sub mnuLeave_Balance_Click()
    strScr_Name = "frmCompany_Info"
    Authority
    If strAllow = "YES" Then
    rLeave_Balance.Show vbModal
    End If
End Sub

Private Sub mnuJournalEntry_Click()
strScr_Name = "frmJournalEntry"
    Authority
    If strAllow = "YES" Then
    frmJournalEntry.Show vbModal
    End If
End Sub

Private Sub mnuLeave_Setup_Click()
    strScr_Name = "frmLeave_Setup"
    Authority
    If strAllow = "YES" Then
    frmLeave_Setup.Show vbModal
    End If
End Sub

Private Sub mnuLPrinting_Click()
strScr_Name = "frmLevelPrint"
    Authority
    If strAllow = "YES" Then
    frmLevelPrint.Show vbModal
    End If
End Sub

Private Sub mnuMamo_Click()
    'strScr_Name = "rMammography"
    'Authority
    'If strAllow = "YES" Then
    rMammography.Show vbModal
    'End If
End Sub

Private Sub mnuMicro_Spec_Click()
     'strScr_Name = "rMicrobiology1"
    'Authority
    'If strAllow = "YES" Then
    rMicrobiology1.Show vbModal
    'End If
End Sub

Private Sub mnuMicroB_Click()
    'strScr_Name = "rMicrobiology"
    'Authority
    'If strAllow = "YES" Then
    rMicrobiology.Show vbModal
    'End If
End Sub

Private Sub mnuNEW_DOC_INFO_Click()
    strScr_Name = "rDoc_New"
        Authority
    If strAllow = "YES" Then
        rDoc_New.Show vbModal
    End If

End Sub
Private Sub mnuPaps_Click()
    'strScr_Name = "rPaps"
    'Authority
    'If strAllow = "YES" Then
    rPaps.Show vbModal
    'End If
End Sub

Private Sub mnuPat_info_Click()
    strScr_Name = "rPat_Info"
    Authority
    If strAllow = "YES" Then
    rPat_Info.Show vbModal
    End If
End Sub

Private Sub mnuPat_Type_Click()
    strScr_Name = "rPat_Type"
        Authority
    If strAllow = "YES" Then
    rPat_Type.Show vbModal
    End If
End Sub

Private Sub MNUPAtient_Info_VAT_Click()
    strScr_Name = "frmPatient_Info_VAT"
        Authority
    If strAllow = "YES" Then
    frmPat_Info_VAT.Show vbModal
    End If

End Sub

Private Sub mnuPatient_Information_Click()
    strScr_Name = "frmPatient_Info"
        Authority
    If strAllow = "YES" Then
    frmPatient_Info.Show vbModal
    End If

End Sub

Private Sub mnuPay_Edit_Click()
    strScr_Name = "frmPay_Edit"
    Authority
    If strAllow = "YES" Then
        frmPay_Edit.Show vbModal
    End If
End Sub

Private Sub mnuPLAccounts_Click()
''strScr_Name = "frmPLAccounts"
''    Authority
''    If strAllow = "YES" Then
''    frmPLAccounts.Show vbModal
''    End If

End Sub

Private Sub mnuRCDetails_Click()
strScr_Name = "frmRefDoc"
    Authority
    If strAllow = "YES" Then
    frmRefDoc.Show vbModal
    End If

End Sub

Private Sub mnuRDPDetails_Click()
strScr_Name = "frmRefDoc"
    Authority
    If strAllow = "YES" Then
    frmRefDoc.Show vbModal
    End If

End Sub

Private Sub mnuGLedger_Click()
strScr_Name = "RptGLedger"
    Authority
    If strAllow = "YES" Then
    RptGLedger.Show vbModal
    End If
End Sub

Private Sub mnuStock_In_Click()
    strScr_Name = "frmStock_IN"
    Authority
    If strAllow = "YES" Then
    frmStock_IN.Show vbModal
    End If
End Sub

Private Sub mnuStock_Status_Click()
    strScr_Name = "rStock_Status"
    Authority
    If strAllow = "YES" Then
    rStock_Status.Show vbModal
    End If
End Sub

Private Sub mnuStool_Click()
    'strScr_Name = "rStool"
    'Authority
    'If strAllow = "YES" Then
    rStool.Show vbModal
    'End If
End Sub

Private Sub mnuStool_Spec_Click()
    'strScr_Name = "frmCompany_Info"
    'Authority
    'If strAllow = "YES" Then
    rStool1.Show vbModal
    'End If
End Sub

Private Sub mnuSup_Info_Click()
    strScr_Name = "frmCompany_Info"
    Authority
    If strAllow = "YES" Then
       frmSup_Info.Show vbModal
    End If
End Sub

Private Sub mnuTBalance_Click()
'strScr_Name = "frmTBalance"
'    Authority
'    If strAllow = "YES" Then
'    frmTBalance.Show vbModal
'    End If

End Sub

Private Sub mnuTest_Information_Click()
    strScr_Name = "frmTest_Info"
        Authority
        
    If strAllow = "YES" Then
    frmTest_Info.Show vbModal
    End If

End Sub

Private Sub mnuTest_Monidy_Click()
 strScr_Name = "frmEdit_TestCode_Type"
 Authority
 If strAllow = "YES" Then
 frmEdit_TestCode_Type.Show vbModal
 End If
End Sub

Private Sub mnuTestResult_Click()
    strScr_Name = "frmTest_Result"
    Authority
    If strAllow = "YES" Then
    frmTest_Result.Show vbModal
    End If
End Sub

Private Sub mnuTM_Click()
    'strScr_Name = "rTumour_Marker"
    'Authority
    'If strAllow = "YES" Then
    rTumour_Marker.Show vbModal
    'End If
End Sub

Private Sub mnuTWStatemetn_Click()
strScr_Name = "RptTWStatement"
    Authority
    If strAllow = "YES" Then
    RptTWStatement.Show vbModal
    End If
End Sub

Private Sub mnuUltra_Click()
    'strScr_Name = "rUltrasonogram"
    'Authority
    'If strAllow = "YES" Then
    rUltrasonogram.Show vbModal
    'End If
End Sub

Private Sub mnuUrine_Click()
    'strScr_Name = "rUrine1"
    'Authority
    'If strAllow = "YES" Then
    rUrine1.Show vbModal
    'End If
End Sub

Private Sub mnuUrine_Spec_Click()

    'strScr_Name = "rUrine"
    'Authority
    'If strAllow = "YES" Then
     rUrine.Show vbModal
    'End If
    
End Sub

Private Sub mnuVAT_Setup_Click()
   ' strScr_Name = "frmVAT_Setup"
   ' Authority
   ' If strAllow = "YES" Then
    frmVAT_Setup.Show vbModal
   ' End If
End Sub

Private Sub mnuVEntry_Click()
strScr_Name = "frmMoneyReceipt"
    Authority
    
    If strAllow = "YES" Then
    frmMoneyReceipt.Show vbModal
    End If
End Sub

Private Sub mnuX_Ray_Click()
    'strScr_Name = "rX_Ray"
    'Authority
    'If strAllow = "YES" Then
    rX_Ray.Show vbModal
    'End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
'    Case 1
'    MsgBox "1"
'    Case 2
'    MsgBox "2"
    
End Select
End Sub

Private Sub Test_Information_Click()
strScr_Name = "frmTest_Info"
        Authority
        
    If strAllow = "YES" Then
    frmTest_Info.Show vbModal
    End If
End Sub
