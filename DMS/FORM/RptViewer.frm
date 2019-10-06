VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form Viewer 
   Caption         =   "Diagnostic management system"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12960
   DrawWidth       =   2
   Icon            =   "RptViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9480
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   11055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15135
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim My_Rst As New ADODB.Recordset
Dim StrSumAdv As Double 'for sum of Product Value to show in word
Dim Report1 As New CrTest_Info ' 'for test_info
Dim Report2 As New CrDoc_Info '
Dim Report3 As New TUMOUR_MARKER '
Dim Report4 As New Hormone '
Dim Report5 As New Histopath 'this Report using for Histopath and PAPS
Dim Report7 As New UltraSono
Dim Report8 As New Echo  'for ECHOCARDIOLOGY
Dim Report9 As New X_RAY
Dim Report10 As New Immunology
Dim Report11 As New BioChamical
Dim Report12 As New Stool
Dim Report14 As New Pat_Info1
Dim Report15 As New Doc_Pay
Dim Report16 As New Doc_Info_New
Dim Report17 As New Pat_Info_VAT
Dim Report18 As New Booth
Dim Report19 As New Hepatities
Dim Report21 As New Urine1
Dim Report22 As New Microbiology
Dim Report23 As New BodyFluid
Dim Report24 As New Drug
Dim Report25 As New Mammography
Dim Report26 As New CT_scan
Dim Report27 As New Endoscopy
Dim Report28 As New Haematology
Dim Report30 As New Daily_Stat2
Dim Report31 As New Adv_Coll
Dim Report32 As New Due_Coll
Dim Report33 As New Due_Pat
Dim Report34 As New Daily_Test
Dim Report35 As New Pat_Type
Dim Report36 As New Pat_Info
Dim Report37 As New Lv_Balance
Dim Report38 As New Stock_Status
Dim Report39 As New Pat_Info2
Dim Report40 As New Doc_Due_Pat
Dim Report41 As New Adv_Coll_All
Dim Report42 As New Due_Coll_All
Dim Report43 As New Urine
Dim Report44 As New Stool1
Dim Report45 As New X_RAY1
Dim Report46 As New Microbiology1
Dim Report47 As New Stock_Balance
'Dim Report48 As New CrysSalesManPerformance

Dim StDate_TM1 As String
Dim EdDate_TM1 As String
Dim StrField4 As String

Private Sub Form_Load()
On Error Resume Next

    CRViewer1.Zoom 100
    Dim StrPat_ID As String
    Dim strM_Code As String
    Dim strS_Code As String
    Dim strSt_date As String
    Dim strEd_date As String
    Dim strRefer_Code As String

    Dim StrStdt As String
    Dim StrSttime As String
    Dim StDate_TM As String
    
    Dim StrEddt As String
    Dim StrEdtime As String
    Dim EdDate_TM As String
             
Select Case CRViewer1_MODE
       Case 1
            strM_Code = RptTest_Info.txtM_Code
            Report1.DiscardSavedData
            If RptTest_Info.Option1 = True Then
               rs.Open "exec rpt_test_info '1',''", strcn.Connection
               Report1.Database.SetDataSource rs
            ElseIf RptTest_Info.Option2 = True Then
               rs.Open "exec rpt_test_info '2','" & RptTest_Info.txtM_Code & "'", strcn.Connection
               Report1.Database.SetDataSource rs
            End If

            CRViewer1.ReportSource = Report1

        Case 2

            strRefer_Code = RptDoctor_Info.txtRefer_Code
            Report2.DiscardSavedData
            If RptDoctor_Info.Option1 = True Then
               rs.Open "exec rpt_doctor_info '1',''", strcn.Connection
               Report2.Database.SetDataSource rs
            ElseIf RptDoctor_Info.Option2 = True Then
               rs.Open "exec rpt_doctor_info '2','" & RptDoctor_Info.txtRefer_Code & "'", strcn.Connection
               Report2.Database.SetDataSource rs
            End If

            CRViewer1.ReportSource = Report2

        Case 3
            StrPat_ID = rTumour_Marker.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rTumour_Marker.txtM_Code
            strS_Code = rTumour_Marker.txtS_Code
            '--------------------------------------------------------------------
            Report3.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report3.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report3.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report3.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report3.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report3.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report3.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report3.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report3.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report3.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report3.FormulaFields.Item(11).text = Chr(34) & "Laboratory Report" & Chr(34)
            Report3.FormulaFields.Item(12).text = Chr(34) & "Name of Tests" & Chr(34)
            Report3.FormulaFields.Item(13).text = Chr(34) & "Results" & Chr(34)
            Report3.FormulaFields.Item(14).text = Chr(34) & "Unit" & Chr(34)
            Report3.FormulaFields.Item(15).text = Chr(34) & "Normal Ranges" & Chr(34)
            Report3.FormulaFields.Item(16).text = Chr(34) & "Checked By" & Chr(34)

            Report3.Text2.SetText Trim(rTumour_Marker.txtUnit.text)

            Call Flush_Doc_Name
            Report3.Text5.SetText StDoc_Name

            Report3.DiscardSavedData
            rs.Open "exec rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection

            Report3.Database.SetDataSource rs
            CRViewer1.ReportSource = Report3
        Case 4
            StrPat_ID = rHormone.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rHormone.txtM_Code
            strS_Code = rHormone.txtS_Code

            '--------------------------------------------------------------------
            Report4.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report4.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report4.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report4.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report4.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report4.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report4.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report4.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report4.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report4.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report4.FormulaFields.Item(11).text = Chr(34) & "Name of Tests" & Chr(34)
            Report4.FormulaFields.Item(12).text = Chr(34) & "Results" & Chr(34)
            Report4.FormulaFields.Item(13).text = Chr(34) & "Unit" & Chr(34)
            Report4.FormulaFields.Item(14).text = Chr(34) & "Normal Ranges" & Chr(34)
            Report4.FormulaFields.Item(15).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report4.Text1.SetText StDoc_Name

            Report4.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report4.Database.SetDataSource rs
            CRViewer1.ReportSource = Report4

        Case 5 'this report is using for histopath and paps
            StrPat_ID = rHistopathology.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rHistopathology.txtM_Code
            strS_Code = rHistopathology.txtS_Code

           '--------------------------------------------------------------------
            Report5.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report5.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report5.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report5.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report5.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report5.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report5.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report5.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report5.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report5.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report5.FormulaFields.Item(11).text = Chr(34) & "Dx" & Chr(34)
            Report5.FormulaFields.Item(12).text = Chr(34) & "Advice" & Chr(34)
            Report5.FormulaFields.Item(13).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report5.Text2.SetText StDoc_Name

            Report5.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report5.Database.SetDataSource rs
            CRViewer1.ReportSource = Report5
        Case 6
            StrPat_ID = rPaps.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rPaps.txtM_Code
            strS_Code = rPaps.txtS_Code

            Report6.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report6.Database.SetDataSource rs
            CRViewer1.ReportSource = Report6
        Case 7
            StrPat_ID = rUltrasonogram.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rUltrasonogram.txtM_Code
            strS_Code = rUltrasonogram.txtS_Code

            Report7.Text1.SetText Trim(rUltrasonogram.txtTest_Name.text)

            Call Flush_Doc_Name
            Report7.Text2.SetText StDoc_Name

            '--------------------------------------------------------------------
            Report7.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report7.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report7.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report7.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report7.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report7.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report7.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report7.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report7.FormulaFields.Item(9).text = Chr(34) & "NATURE OF EXAM" & Chr(34)
            Report7.FormulaFields.Item(10).text = Chr(34) & "ULTRSONOGRAM REPORT" & Chr(34)
            Report7.FormulaFields.Item(11).text = Chr(34) & "Checked By" & Chr(34)

            Report7.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report7.Database.SetDataSource rs
            CRViewer1.ReportSource = Report7
        Case 8
            StrPat_ID = rEchocardiography.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rEchocardiography.txtM_Code
            strS_Code = rEchocardiography.txtS_Code

           '--------------------------------------------------------------------
            Report8.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report8.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report8.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report8.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report8.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report8.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report8.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report8.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report8.FormulaFields.Item(9).text = Chr(34) & "MEASUREMENTS" & Chr(34)
            Report8.FormulaFields.Item(10).text = Chr(34) & "DESCRIPTION" & Chr(34)
            Report8.FormulaFields.Item(11).text = Chr(34) & "IMPRESSION" & Chr(34)
            Report8.FormulaFields.Item(12).text = Chr(34) & "Checked By" & Chr(34)

            Report8.Text1.SetText rEchocardiography.txtTest_Result1.text

            Call Flush_Doc_Name
            Report8.Text1.SetText StDoc_Name


            Report8.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report8.Database.SetDataSource rs
            CRViewer1.ReportSource = Report8
        Case 9
            StrPat_ID = rX_Ray.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rX_Ray.txtM_Code
            strS_Code = rX_Ray.txtS_Code

            '--------------------------------------------------------------------
            Report9.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report9.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report9.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report9.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report9.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report9.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report9.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report9.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report9.FormulaFields.Item(9).text = Chr(34) & "X-RAY REPORT" & Chr(34)
            Report9.FormulaFields.Item(10).text = Chr(34) & "Comment" & Chr(34)
            Report9.FormulaFields.Item(11).text = Chr(34) & "Advice" & Chr(34)
            Report9.FormulaFields.Item(12).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report9.Text1.SetText StDoc_Name

            Report9.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            '--------------------------
            Dim StrField44 As String
            If rs.EOF = False Then
                StrField44 = rs!Field4
            End If

            '--------------------------

            Report9.Database.SetDataSource rs
            CRViewer1.ReportSource = Report9
        Case 10
            StrPat_ID = rImmunology.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rImmunology.txtM_Code
            strS_Code = rImmunology.txtS_Code

            Report10.Text23.SetText Trim(rImmunology.txtNote.text)

            Call Flush_Doc_Name
            Report10.Text1.SetText StDoc_Name

            Report10.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report10.Database.SetDataSource rs
            CRViewer1.ReportSource = Report10

        Case 11

            StrPat_ID = rBio_Chamical.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rBio_Chamical.txtM_Code
            strS_Code = rBio_Chamical.txtS_Code
            '--------------------------------------------------------------------
            Report11.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report11.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report11.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report11.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report11.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report11.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report11.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report11.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report11.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report11.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report11.FormulaFields.Item(11).text = Chr(34) & "Laboratory Report" & Chr(34)
            Report11.FormulaFields.Item(12).text = Chr(34) & "Name of Tests" & Chr(34)
            Report11.FormulaFields.Item(13).text = Chr(34) & "Results" & Chr(34)
            Report11.FormulaFields.Item(14).text = Chr(34) & "Unit" & Chr(34)
            Report11.FormulaFields.Item(15).text = Chr(34) & "Normal Ranges" & Chr(34)
            Report11.FormulaFields.Item(16).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report11.Text1.SetText StDoc_Name


            Report11.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report11.Database.SetDataSource rs
            CRViewer1.ReportSource = Report11

        Case 12
            StrPat_ID = rStool.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rStool.txtM_Code
            strS_Code = rStool.txtS_Code
            StrPat_ID = rStool.txtPat_ID
            strM_Code = rStool.txtM_Code
            strS_Code = rStool.txtS_Code

            '--------------------------------------------------------------------
            Report12.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report12.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report12.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report12.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report12.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report12.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report12.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report12.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report12.FormulaFields.Item(9).text = Chr(34) & "STOOL EXAMINATION REPORT" & Chr(34)
            Report12.FormulaFields.Item(10).text = Chr(34) & "Checked By" & Chr(34)


            Call Flush_Doc_Name
            Report12.Text1.SetText StDoc_Name


            Report12.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report12.Database.SetDataSource rs
            CRViewer1.ReportSource = Report12

        Case 13
            StrPat_ID = rStool1.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rStool1.txtM_Code
            strS_Code = rStool1.txtS_Code

            Report13.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report13.Database.SetDataSource rs
            CRViewer1.ReportSource = Report13
       Case 14
'            If frmPatient_Info.txtPat_ID = "" Then
'            StrPat_ID = StPat_ID
'            Else
'            StrPat_ID = frmPatient_Info.txtPat_ID
'            End If
'
'            Report14.DiscardSavedData
'            rs.Open "exec Rpt_pat_info '" & StrPat_ID & "'", strcn.Connection
'            Report14.Database.SetDataSource rs
'            Report14.Text29.SetText frmPatient_Info.txtRefer_Code.text
'            CRViewer1.ReportSource = Report14
'            Printer.PaperSize = 9
'           Printer.oriantation = 2
'            Report.PaperSize = crPaperLetter
'            CRViewer1.PaperSize = crPaperEnvelope12
'Printer.Orientation = vbPRORLandscape


       Case 15

             StrStdt = Trim(Format(rDoc_Pay.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rDoc_Pay.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

            StrEddt = Trim(Format(rDoc_Pay.edDt, "yyyy-mm-dd"))
            StrEdtime = Trim(Format(rDoc_Pay.edDT_TM, "hh:mm"))
            EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
            Report15.FormulaFields.Item(17).text = Chr(34) & Format(StDate_TM, "dd/mm/yyyy") & Chr(34)
            Report15.FormulaFields.Item(18).text = Chr(34) & Format(EdDate_TM, "dd/mm/yyyy") & Chr(34)

            strRefer_Code = rDoc_Pay.txtRefer_Code
            strSt_date = Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM")
            strEd_date = Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM")


            If rDoc_Pay.Chk_Hide_Tot.value = "1" Then
                Report15.FormulaFields.Item(22).text = "1"
            End If

            If rDoc_Pay.Chk_Hide_Tot.value = "0" Then
                Report15.FormulaFields.Item(22).text = "0"
            End If

            Report15.DiscardSavedData

            rs.Open "exec Rpr_Doc_Pay3 1,'" & strRefer_Code & "'", strcn.Connection

            Report15.Database.SetDataSource rs
            CRViewer1.ReportSource = Report15

        Case 16
            Dim strSt_date1 As String
            Dim strEd_date1 As String
            Dim StrDoc_Name As String

            StrDoc_Name = Trim(rDoc_New.CombDoc_Name)

            strSt_date1 = Format(rDoc_New.StDate, "yyyy-mm-dd")
            strEd_date1 = Format(rDoc_New.EdDate, "yyyy-mm-dd")

            Report16.FormulaFields.Item(1).text = Chr(34) & rDoc_New.StDate & Chr(34)
            Report16.FormulaFields.Item(2).text = Chr(34) & rDoc_New.EdDate & Chr(34)

            Report16.DiscardSavedData
            If rDoc_New.Option1 = True Then

               rs.Open "exec Rpt_Doctor_Info_New '1','','" + strSt_date1 + "','" + strEd_date1 + "'", strcn.Connection
               Report16.Database.SetDataSource rs
            Else
               rs.Open "exec Rpt_Doctor_Info_New '2','" + StrDoc_Name + "','" + strSt_date1 + "','" + strEd_date1 + "'", strcn.Connection
               Report16.Database.SetDataSource rs
            End If
            CRViewer1.ReportSource = Report16

       Case 17
            '+++++++++++++Starting Date++++++++++++
             StrStdt = Trim(Format(frmPat_Info_VAT.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(frmPat_Info_VAT.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(frmPat_Info_VAT.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(frmPat_Info_VAT.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            Report17.DiscardSavedData
            rs.Open "exec Rpt_VAT_Pat '" + StDate_TM + "','" + EdDate_TM + "'", strcn.Connection
            Report17.Database.SetDataSource rs
            CRViewer1.ReportSource = Report17

        Case 18

            Dim StrUid As String
            StrUid = rBooth_User_Info.txtUID


            '++++++for Starting Date and Time++++++++++++++
             StrStdt = Trim(Format(rBooth_User_Info.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rBooth_User_Info.stDT_TM, "hh:mm AM/PM"))
             StDate_TM1 = StrStdt + Space(3) + StrSttime
'             StrStdt = Trim(Format(rBooth_User_Info.stDt, "dd-MM-yyyy"))
'             StrSttime = Trim(Format(rBooth_User_Info.stDT_TM, "hh:mm AM/PM"))
'
'             StDate_TM1 = StrStdt + Space(3) + StrSttime
'
             StDate_TM = StrStdt

            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++
             StrEddt = Trim(Format(rBooth_User_Info.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rBooth_User_Info.edDT_TM, "hh:mm AM/PM"))
             EdDate_TM1 = StrEddt + Space(3) + StrEdtime
'             StrEddt = Trim(Format(rBooth_User_Info.edDt, "dd-MM-yyyy"))
'             StrEdtime = Trim(Format(rBooth_User_Info.edDT_TM, "hh:mm AM/PM"))
'             'EdDate_TM = StrEddt + Space(1) + StrEdtime
'             EdDate_TM1 = StrEddt + Space(3) + StrEdtime
'
             EdDate_TM = StrEddt


            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '///////for use CONVERT FUNCTION////////////////////
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            My_Rst.Open "select sum(adv) as SumAdv from pat_info_sub2 where uid='" + StrUid + "' and dt between '" + Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "' and '" + Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "'", con
            If IsNull(My_Rst!SumAdv) = False Then
                StrSumAdv = My_Rst!SumAdv
                Report18.FormulaFields.Item(3).text = Chr(34) & ConvertX(StrSumAdv) & Chr(34)

            End If
            con.Close
            '/////////////////////////////////////////////////

            Report18.FormulaFields.Item(1).text = Chr(34) & StDate_TM1 & Chr(34)
            Report18.FormulaFields.Item(2).text = Chr(34) & EdDate_TM1 & Chr(34)
            Report18.FormulaFields.Item(6).text = Chr(34) & rBooth_User_Info.txtBooth & Chr(34)

            Report18.DiscardSavedData
            If rBooth_User_Info.Option1.value = True Then
                Report18.FormulaFields.Item(5).text = Chr(34) & 0 & Chr(34)
                rs.Open "exec rpt_booth '" & rBooth_User_Info.txtBooth.text & "','" & Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "'", strcn.Connection
            End If

            If rBooth_User_Info.Option2.value = True Then
            Report18.FormulaFields.Item(5).text = Chr(34) & 1 & Chr(34)
                rs.Open "exec rpt_booth1 '" & rBooth_User_Info.txtUID & "','" & rBooth_User_Info.txtBooth & "','" + Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "','" + Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "'", strcn.Connection
            End If

            Report18.Database.SetDataSource rs
            CRViewer1.ReportSource = Report18

         Case 19

            StrPat_ID = rHepatitis.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rHepatitis.txtM_Code
            strS_Code = rHepatitis.txtS_Code

            '--------------------------------------------------------------------
            Report19.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report19.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report19.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report19.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report19.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report19.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report19.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report19.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report19.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report19.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report19.FormulaFields.Item(11).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report19.Text1.SetText StDoc_Name

            Report19.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report19.Database.SetDataSource rs
            CRViewer1.ReportSource = Report19
        Case 20
            StrPat_ID = rUrine.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rUrine.txtM_Code
            strS_Code = rUrine.txtS_Code

            Report20.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report20.Database.SetDataSource rs
            CRViewer1.ReportSource = Report20
        Case 21

            StrPat_ID = rUrine1.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rUrine1.txtM_Code
            strS_Code = rUrine1.txtS_Code
            '--------------------------------------------------------------------
            Report21.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report21.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report21.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report21.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report21.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report21.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report21.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report21.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report21.FormulaFields.Item(9).text = Chr(34) & "URINE EXAMINATION REPORT" & Chr(34)
            Report21.FormulaFields.Item(10).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report21.Text1.SetText StDoc_Name

            Report21.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report21.Database.SetDataSource rs
            CRViewer1.ReportSource = Report21

        Case 22

            StrPat_ID = rMicrobiology.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rMicrobiology.txtM_Code
            strS_Code = rMicrobiology.txtS_Code

           '--------------------------------------------------------------------
            Report22.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report22.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report22.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report22.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report22.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report22.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report22.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report22.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report22.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report22.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report22.FormulaFields.Item(11).text = Chr(34) & "1. Organism Isolated" & Chr(34)
            Report22.FormulaFields.Item(12).text = Chr(34) & "2. Sensitivity Test" & Chr(34)
            Report22.FormulaFields.Item(13).text = Chr(34) & "Checked By" & Chr(34)

            Report22.Text2.SetText Trim(rMicrobiology.txtTest_Name.text)
            Report22.Text3.SetText Trim(rMicrobiology.txtTest_Result.text)

            Call Flush_Doc_Name
            Report22.Text6.SetText StDoc_Name

            Report22.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report22.Database.SetDataSource rs
            CRViewer1.ReportSource = Report22

        Case 23

            StrPat_ID = rBody_Fluid.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rBody_Fluid.txtM_Code
            strS_Code = rBody_Fluid.txtS_Code

           '--------------------------------------------------------------------
            Report23.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report23.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report23.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report23.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report23.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report23.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report23.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report23.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report23.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report23.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report23.FormulaFields.Item(11).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report23.Text1.SetText StDoc_Name


            Report23.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report23.Database.SetDataSource rs
            CRViewer1.ReportSource = Report23
        Case 24

            StrPat_ID = rDrug.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rDrug.txtM_Code
            strS_Code = rDrug.txtS_Code

            '--------------------------------------------------------------------
            Report24.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report24.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report24.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report24.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report24.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report24.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report24.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report24.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report24.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report24.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report24.FormulaFields.Item(11).text = Chr(34) & "Name of Tests" & Chr(34)
            Report24.FormulaFields.Item(12).text = Chr(34) & "Results" & Chr(34)
            Report24.FormulaFields.Item(13).text = Chr(34) & "Unit" & Chr(34)
            Report24.FormulaFields.Item(14).text = Chr(34) & "Normal Ranges" & Chr(34)
            Report24.FormulaFields.Item(15).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report24.Text1.SetText StDoc_Name

            Report24.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report24.Database.SetDataSource rs
            CRViewer1.ReportSource = Report24
    Case 25
            StrPat_ID = rMammography.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rMammography.txtM_Code
            strS_Code = rMammography.txtS_Code

            '--------------------------------------------------------------------
            Report25.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report25.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report25.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report25.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report25.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report25.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report25.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report25.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report25.FormulaFields.Item(9).text = Chr(34) & "MEMMOGRAPHY REPORT" & Chr(34)
            Report25.FormulaFields.Item(10).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report25.Text1.SetText StDoc_Name

            Report25.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report25.Database.SetDataSource rs
            CRViewer1.ReportSource = Report25
        Case 26
            StrPat_ID = rCT_SCAN.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rCT_SCAN.txtM_Code
            strS_Code = rCT_SCAN.txtS_Code

            '--------------------------------------------------------------------
            Report26.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report26.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report26.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report26.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report26.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report26.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report26.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report26.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report26.FormulaFields.Item(9).text = Chr(34) & "C.T. SCAN REPORT" & Chr(34)
            Report26.FormulaFields.Item(10).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report26.Text2.SetText StDoc_Name


            Report26.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report26.Database.SetDataSource rs
            CRViewer1.ReportSource = Report26
        Case 27

            StrPat_ID = rEndoscopy.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rEndoscopy.txtM_Code
            strS_Code = rEndoscopy.txtS_Code

            '--------------------------------------------------------------------
            Report27.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report27.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report27.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report27.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report27.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report27.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report27.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report27.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report27.FormulaFields.Item(9).text = Chr(34) & "ENDOSCOPY REPORT" & Chr(34)
            Report27.FormulaFields.Item(10).text = Chr(34) & "IMPRESSION" & Chr(34)
            Report27.FormulaFields.Item(11).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report27.Text1.SetText StDoc_Name


            Report27.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report27.Database.SetDataSource rs
            CRViewer1.ReportSource = Report27
        Case 28
            StrPat_ID = rHaematology.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rHaematology.txtM_Code
            strS_Code = rHaematology.txtS_Code

            '--------------------------------------------------------------------
            Report28.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report28.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report28.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report28.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report28.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report28.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report28.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report28.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report28.FormulaFields.Item(9).text = Chr(34) & "Haematological Analysis Report" & Chr(34)
            Report28.FormulaFields.Item(10).text = Chr(34) & "Tests" & Chr(34)
            Report28.FormulaFields.Item(11).text = Chr(34) & "Results" & Chr(34)
            Report28.FormulaFields.Item(12).text = Chr(34) & "Normal Ranges" & Chr(34)
            Report28.FormulaFields.Item(13).text = Chr(34) & "Checked By" & Chr(34)

            If rHaematology.txtS_Code = "23" Or rHaematology.txtS_Code = "24" Then
            Report28.FormulaFields.Item(14).text = Chr(34) & "Test are carried out by SYSMEX KX-21" & Chr(34)
            End If

            Report28.Text2.SetText Trim(rHaematology.txtTest_Result1.text)

            Call Flush_Doc_Name
            Report28.Text4.SetText StDoc_Name

            Report28.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report28.Database.SetDataSource rs
            CRViewer1.ReportSource = Report28
        Case 29 'this report is using for histopath and paps
            StrPat_ID = rPaps.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rPaps.txtM_Code
            strS_Code = rPaps.txtS_Code

           '--------------------------------------------------------------------
            Report5.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report5.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report5.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report5.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report5.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report5.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report5.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report5.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report5.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report5.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report5.FormulaFields.Item(11).text = Chr(34) & "Dx" & Chr(34)
            Report5.FormulaFields.Item(12).text = Chr(34) & "Advice" & Chr(34)
            Report5.FormulaFields.Item(13).text = Chr(34) & "Checked By" & Chr(34)

            Report5.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report5.Database.SetDataSource rs
            CRViewer1.ReportSource = Report5
        Case 30
             StrStdt = Trim(Format(rDaily_Statement.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rDaily_Statement.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rDaily_Statement.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rDaily_Statement.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '----end--------------------------

            '///////for use CONVERT FUNCTION////////////////////
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con
'            con.Close
            '////////////////////////////////////////////////

            Report30.FormulaFields.Item(1).text = Chr(34) & Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
            Report30.FormulaFields.Item(2).text = Chr(34) & Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
            Report30.FormulaFields.Item(5).text = Chr(34) & Format(rDaily_Statement.stDt, "YYYY") & Chr(34)

            Report30.DiscardSavedData

               rs.Open "exec Daily_Stat2 '" & StDate_TM & "','" & EdDate_TM & "'", strcn.Connection
               Report30.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report30
        Case 31
             StrStdt = Trim(Format(rAdv_Coll.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rAdv_Coll.stDT_TM, "hh:mm AM/PM"))
             StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rAdv_Coll.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rAdv_Coll.edDT_TM, "hh:mm AM/PM"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++


            '///////for use CONVERT FUNCTION////////////////////
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con
'            con.Close
            '////////////////////////////////////////////////

            Report31.FormulaFields.Item(1).text = Chr(34) & Format(StrStdt, "dd/mm/yyyy") & Chr(34)
            Report31.FormulaFields.Item(2).text = Chr(34) & Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
            Report31.FormulaFields.Item(3).text = Chr(34) & StrSttime & Chr(34)
            Report31.FormulaFields.Item(4).text = Chr(34) & StrEdtime & Chr(34)

            Report31.DiscardSavedData

               rs.Open "exec Advance_Coll 1,'" + rAdv_Coll.txtU_ID + "','" + StDate_TM + "','" + EdDate_TM + "'", strcn.Connection

               Report31.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report31
    Case 32
             StrStdt = Trim(Format(rAdv_Coll.stDt, "dd-mm-yyyy"))
             StrSttime = Trim(Format(rAdv_Coll.stDT_TM, "hh:mm AM/PM"))
             StDate_TM = StrStdt + Space(1) + StrSttime

             StDate_TM1 = StrStdt + Space(3) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rAdv_Coll.edDt, "dd-mm-yyyy"))
             StrEdtime = Trim(Format(rAdv_Coll.edDT_TM, "hh:mm AM/PM"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime

             EdDate_TM1 = StrEddt + Space(3) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '///////for use CONVERT FUNCTION////////////////////
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            Report32.FormulaFields.Item(1).text = Chr(34) & StDate_TM1 & Chr(34)

            Report32.FormulaFields.Item(2).text = Chr(34) & EdDate_TM1 & Chr(34)
            Report32.DiscardSavedData

            rs.Open "exec due_Coll '" + rAdv_Coll.txtU_ID.text + "','" + Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "','" + Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "'", strcn.Connection
'               Debug.Print "exec Advance_Coll '" + StDate_TM + "','" + EdDate_TM + "'"


            Report32.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report32
    Case 33

             StrStdt = Trim(Format(rAdv_Coll.stDt, "dd-mm-yyyy"))
             StrSttime = Trim(Format(rAdv_Coll.stDT_TM, "hh:mm AM/PM"))
             StDate_TM = StrStdt + Space(1) + StrSttime
             StDate_TM1 = StrStdt + Space(3) + StrSttime

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rAdv_Coll.edDt, "dd-mm-yyyy"))
             StrEdtime = Trim(Format(rAdv_Coll.edDT_TM, "hh:mm AM/PM"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
             EdDate_TM1 = StrEddt + Space(3) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '///////for use CONVERT FUNCTION////////////////////
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con
            Report33.FormulaFields.Item(1).text = Chr(34) & StDate_TM1 & Chr(34)
            Report33.FormulaFields.Item(2).text = Chr(34) & EdDate_TM1 & Chr(34)
            Report33.DiscardSavedData

            rs.Open "select*from pat_due_coll", strcn.Connection

            Report33.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report33
      Case 34
             StrStdt = Trim(Format(rDaily_Test.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rDaily_Test.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime
             StDate_TM1 = StrStdt + Space(3) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rDaily_Test.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rDaily_Test.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
             EdDate_TM1 = StrEddt + Space(3) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            Report34.FormulaFields.Item(1).text = Chr(34) & Format(StDate_TM1, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
            Report34.FormulaFields.Item(2).text = Chr(34) & Format(EdDate_TM1, "dd/mm/yyyy hh:mm AMPM") & Chr(34)

            Report34.DiscardSavedData
               If rDaily_Test.Option1 = True Then
               rs.Open "exec rptTest_State 1,'','" & Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "'", strcn.Connection
               End If

               If rDaily_Test.Option2 = True Then
               rs.Open "exec rptTest_State 2,'" & Trim(rDaily_Test.txtM_Code.text) & "','" & Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "'", strcn.Connection
               End If

            Report34.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report34
      Case 35
             StrStdt = Trim(Format(rPat_Type.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rPat_Type.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rPat_Type.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rPat_Type.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            Report35.FormulaFields.Item(1).text = Chr(34) & Format(StDate_TM, "dd-mm-yyyy hh:mm AMPM") & Chr(34)
            Report35.FormulaFields.Item(2).text = Chr(34) & Format(EdDate_TM, "dd-mm-yyyy hh:mm AMPM") & Chr(34)
            Report35.DiscardSavedData

               If rPat_Type.Option1.value = True Then
               Report35.FormulaFields.Item(3).text = Chr(34) & " Outside " & " Patient " & " List " & Chr(34)
               rs.Open "exec pat_type '1','" + StDate_TM + "','" + EdDate_TM + "'", strcn.Connection
               End If

               If rPat_Type.Option2.value = True Then
               Report35.FormulaFields.Item(3).text = Chr(34) & "General " & " Patient " & " List " & Chr(34)
               rs.Open "exec pat_type '0','" + StDate_TM + "','" + EdDate_TM + "'", strcn.Connection
               End If

            Report35.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report35
        
        Case 43

            StrPat_ID = rUrine.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rUrine.txtM_Code
            strS_Code = rUrine.txtS_Code
            '--------------------------------------------------------------------
            Report43.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report43.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report43.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report43.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report43.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report43.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report43.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report43.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report43.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report43.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report43.FormulaFields.Item(11).text = Chr(34) & "Laboratory Report" & Chr(34)
            Report43.FormulaFields.Item(12).text = Chr(34) & "Name of Tests" & Chr(34)
            Report43.FormulaFields.Item(13).text = Chr(34) & "Results" & Chr(34)
            Report43.FormulaFields.Item(14).text = Chr(34) & "Unit" & Chr(34)
            Report43.FormulaFields.Item(15).text = Chr(34) & "Normal Ranges" & Chr(34)
            Report43.FormulaFields.Item(16).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report43.Text1.SetText StDoc_Name

            Report43.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report43.Database.SetDataSource rs
            
            CRViewer1.ReportSource = Report43
            
    Case 49

             StrStdt = Trim(Format(rptMExecutive.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rptMExecutive.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rptMExecutive.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rptMExecutive.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
             Report49.FormulaFields.Item(1).text = Chr(34) & Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
             Report49.FormulaFields.Item(2).text = Chr(34) & Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)

             strRefer_Code = rptMExecutive.txtEmp_ID
             strSt_date = Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM")
             strEd_date = Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM")

             Report49.DiscardSavedData

             rs.Open "exec Emp_Performance '" + strRefer_Code + "','" + Format(StDate_TM, "yyyy-mm-dd hh:mm AMPM") + "','" + Format(EdDate_TM, "yyyy-mm-dd hh:mm AMPM") + "'", strcn.Connection

             Report49.Database.SetDataSource rs
             CRViewer1.ReportSource = Report49
    
    End Select
    
    Screen.MousePointer = vbHourglass
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs.Close
End Sub

Private Sub InsDaily_State()
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec Daily_Stat2 '" + u_id + "'"
    cmd.Execute
    con.Close
End Sub

Private Sub Flush_Field()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Rpt 1,'" + rX_Ray.txtPat_ID + "','" + rX_Ray.txtM_Code + "','" + rX_Ray.txtS_Code + "'", con
    If My_Rst.EOF = False Then
        StrField4 = My_Rst!Field4
    End If
    con.Close
End Sub



