VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form Viewer1 
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "RptViewer1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
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
Attribute VB_Name = "Viewer1"
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
    Dim StrEmp_Code As String

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
            Report9.FormulaFields.Item(10).text = Chr(34) & "IMPRESSION" & Chr(34)
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

             StrStdt = Trim(Format(rBooth_User_Info.stDt, "dd-mm-yyyy"))
             StrSttime = Trim(Format(rBooth_User_Info.stDT_TM, "hh:mm AM/PM"))

             StDate_TM1 = StrStdt + Space(3) + StrSttime

             StDate_TM = StrStdt

            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rBooth_User_Info.edDt, "dd-mm-yyyy"))
             StrEdtime = Trim(Format(rBooth_User_Info.edDT_TM, "hh:mm AM/PM"))
             'EdDate_TM = StrEddt + Space(1) + StrEdtime
             EdDate_TM1 = StrEddt + Space(3) + StrEdtime

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
             StrStdt = Trim(Format(rAdv_Coll.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rAdv_Coll.stDT_TM, "hh:mm AM/PM"))
             StDate_TM = StrStdt + Space(1) + StrSttime

             StDate_TM1 = StrStdt + Space(3) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rAdv_Coll.edDt, "yyyy-mm-dd"))
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

             StrStdt = Trim(Format(rAdv_Coll.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rAdv_Coll.stDT_TM, "hh:mm AM/PM"))
             StDate_TM = StrStdt + Space(1) + StrSttime
             StDate_TM1 = StrStdt + Space(3) + StrSttime

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rAdv_Coll.edDt, "yyyy-mm-dd"))
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

            Report35.FormulaFields.Item(1).text = Chr(34) & Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
            Report35.FormulaFields.Item(2).text = Chr(34) & Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
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
        Case 36
        
        Dim StrEid As String
            StrEid = rPat_Info.txtEmp_ID

            '++++++for Starting Date and Time++++++++++++++

'             StrStdt = Trim(Format(rPat_Info.stDt, "dd-mm-yyyy"))
             StrStdt = Trim(Format(rPat_Info.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rPat_Info.stDT_TM, "hh:mm AM/PM"))

             StDate_TM1 = StrStdt + Space(3) + StrSttime

             StDate_TM = StrStdt

            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rPat_Info.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rPat_Info.edDT_TM, "hh:mm AM/PM"))
             EdDate_TM1 = StrEddt + Space(3) + StrEdtime

             EdDate_TM = StrEddt
    
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            My_Rst.Open "select sum(adv) as SumAdv from pat_info_sub2 where uid='" + StrUid + "' and dt between '" + Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "' and '" + Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "'", con
            If IsNull(My_Rst!SumAdv) = False Then
            StrSumAdv = My_Rst!SumAdv
            Report36.FormulaFields.Item(4).text = Chr(34) & ConvertX(StrSumAdv) & Chr(34)

            End If
            con.Close
            '/////////////////////////////////////////////////

            Report36.FormulaFields.Item(1).text = Chr(34) & StDate_TM1 & Chr(34)
            Report36.FormulaFields.Item(2).text = Chr(34) & EdDate_TM1 & Chr(34)
            Report36.FormulaFields.Item(3).text = Chr(34) & rPat_Info.txtEmp_Name & Chr(34)

            Report36.DiscardSavedData
'            If rBooth_User_Info.Option1.value = True Then
'                Report36.FormulaFields.Item(5).text = Chr(34) & 0 & Chr(34)
                rs.Open "exec Rpt_Emp '" & rPat_Info.txtEmp_ID.text & "','" & Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "'", strcn.Connection

            Report36.Database.SetDataSource rs
            CRViewer1.ReportSource = Report36

        Case 37
            Report37.DiscardSavedData
               rs.Open "exec leave_balance 1,'" + rLeave_Balance.txtEmp_ID.text + "'", strcn.Connection

               Report37.Database.SetDataSource rs

            CRViewer1.ReportSource = Report37
        Case 38
            '++++++for Ending Date and Time++++++++++++++
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            Report38.DiscardSavedData

            rs.Open "exec pro_stock_det", strcn.Connection

            Report38.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report38

        Case 39
            Dim StrCid As String
            StrCid = RptConsultant.txtCons_Code

            '++++++for Starting Date and Time++++++++++++++

'             StrStdt = Trim(Format(rPat_Info.stDt, "dd-mm-yyyy"))
             StrStdt = Trim(Format(RptConsultant.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(RptConsultant.stDT_TM, "hh:mm AM/PM"))

             StDate_TM1 = StrStdt + Space(3) + StrSttime

             StDate_TM = StrStdt

            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(RptConsultant.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(RptConsultant.edDT_TM, "hh:mm AM/PM"))
             EdDate_TM1 = StrEddt + Space(3) + StrEdtime

             EdDate_TM = StrEddt
    
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            My_Rst.Open "select sum(adv) as SumAdv from pat_info_sub2 where uid='" + StrUid + "' and dt between '" + Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "' and '" + Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") + "'", con
            If IsNull(My_Rst!SumAdv) = False Then
            StrSumAdv = My_Rst!SumAdv
            Report39.FormulaFields.Item(4).text = Chr(34) & ConvertX(StrSumAdv) & Chr(34)

            End If
            con.Close
            '/////////////////////////////////////////////////

            Report39.FormulaFields.Item(1).text = Chr(34) & StDate_TM1 & Chr(34)
            Report39.FormulaFields.Item(2).text = Chr(34) & EdDate_TM1 & Chr(34)
            Report39.FormulaFields.Item(3).text = Chr(34) & RptConsultant.txtDoc_Name & Chr(34)

            Report39.DiscardSavedData
'            If rBooth_User_Info.Option1.value = True Then
'                Report36.FormulaFields.Item(5).text = Chr(34) & 0 & Chr(34)
                rs.Open "exec Rpt_Consultant '" & RptConsultant.txtCons_Code.text & "','" & Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "'", strcn.Connection
                
'                rs.Open "exec Rpt_Consultant2 '" & Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "'", strcn.Connection


            Report39.Database.SetDataSource rs
            CRViewer1.ReportSource = Report39

      Case 40

             StrStdt = Trim(Format(rDoc_Due_Pat.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rDoc_Due_Pat.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rDoc_Due_Pat.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rDoc_Due_Pat.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++
             Report40.FormulaFields.Item(1).text = Chr(34) & Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
             Report40.FormulaFields.Item(2).text = Chr(34) & Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)

             strRefer_Code = rDoc_Due_Pat.txtRefer_Code
             strSt_date = Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM")
             strEd_date = Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM")

             Report40.DiscardSavedData

             rs.Open "exec due_doc_pat '" + strRefer_Code + "','" + Format(StDate_TM, "yyyy-mm-dd hh:mm AMPM") + "','" + Format(EdDate_TM, "yyyy-mm-dd hh:mm AMPM") + "'", strcn.Connection

             Report40.Database.SetDataSource rs
             CRViewer1.ReportSource = Report40

       Case 41

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
            con.ConnectionTimeout = 0
            con.Open
            Set cmd.ActiveConnection = con
            cmd.CommandTimeout = 0

            '////////////////////////////////////////////////

            Report41.FormulaFields.Item(1).text = Chr(34) & Format(StrStdt, "dd-mm-yyyy") & Chr(34)
            Report41.FormulaFields.Item(2).text = Chr(34) & Format(EdDate_TM, "dd-mm-yyyy hh:mm AMPM") & Chr(34)
            Report41.FormulaFields.Item(3).text = Chr(34) & StrSttime & Chr(34)
            Report41.FormulaFields.Item(4).text = Chr(34) & StrEdtime & Chr(34)

            Report41.DiscardSavedData

               rs.Open "exec Advance_Coll1 1,'" + StDate_TM + "','" + EdDate_TM + "'", strcn.Connection
               Report41.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report41

        Case 42

             StrStdt = Trim(Format(rAdv_Coll.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(rAdv_Coll.stDT_TM, "hh:mm AM/PM"))
             StDate_TM = StrStdt + Space(1) + StrSttime

             StDate_TM1 = StrStdt + Space(3) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(rAdv_Coll.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(rAdv_Coll.edDT_TM, "hh:mm AM/PM"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime

             EdDate_TM1 = StrEddt + Space(3) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '///////for use CONVERT FUNCTION////////////////////
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            Report42.FormulaFields.Item(1).text = Chr(34) & StDate_TM1 & Chr(34)

            Report42.FormulaFields.Item(2).text = Chr(34) & EdDate_TM1 & Chr(34)
            Report42.DiscardSavedData

            rs.Open "exec due_Coll_all 1,'" & Format(StDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "','" & Format(EdDate_TM, "yyyy-mm-dd hh:mm AM/PM") & "'", strcn.Connection

            Report42.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report42

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
        Case 44

            StrPat_ID = rStool1.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rStool1.txtM_Code
            strS_Code = rStool1.txtS_Code
            '--------------------------------------------------------------------
            Report44.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report44.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report44.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report44.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report44.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report44.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report44.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report44.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report44.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report44.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report44.FormulaFields.Item(11).text = Chr(34) & "Laboratory Report" & Chr(34)
            Report44.FormulaFields.Item(12).text = Chr(34) & "Name of Tests" & Chr(34)
            Report44.FormulaFields.Item(13).text = Chr(34) & "Results" & Chr(34)
            Report44.FormulaFields.Item(14).text = Chr(34) & "Unit" & Chr(34)
            Report44.FormulaFields.Item(15).text = Chr(34) & "Normal Ranges" & Chr(34)
            Report44.FormulaFields.Item(16).text = Chr(34) & "Checked By" & Chr(34)


            Call Flush_Doc_Name
            Report44.Text1.SetText StDoc_Name


            Report44.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report44.Database.SetDataSource rs
            CRViewer1.ReportSource = Report44
        Case 45
            StrPat_ID = rX_Ray.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rX_Ray.txtM_Code
            strS_Code = rX_Ray.txtS_Code

            '--------------------------------------------------------------------
            Report45.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report45.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report45.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report45.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report45.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report45.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report45.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report45.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report45.FormulaFields.Item(9).text = Chr(34) & "X-RAY REPORT" & Chr(34)
            Report45.FormulaFields.Item(10).text = Chr(34) & "IMPRESSION" & Chr(34)
            Report45.FormulaFields.Item(11).text = Chr(34) & "Advice" & Chr(34)
            Report45.FormulaFields.Item(12).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report45.Text2.SetText StDoc_Name

            Flush_Field
            Report45.Text1.SetText StrField4

            Report45.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report45.Database.SetDataSource rs
            CRViewer1.ReportSource = Report45
        Case 46

            StrPat_ID = rMicrobiology1.txtPat_ID
            StrPat_ID_R = StrPat_ID

            strM_Code = rMicrobiology1.txtM_Code
            strS_Code = rMicrobiology1.txtS_Code

           '--------------------------------------------------------------------
            Report46.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report46.FormulaFields.Item(2).text = Chr(34) & "Patient's ID" & Chr(34)
            Report46.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report46.FormulaFields.Item(4).text = Chr(34) & "Delivery Date" & Chr(34)
            Report46.FormulaFields.Item(5).text = Chr(34) & "Patient's Name" & Chr(34)
            Report46.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report46.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report46.FormulaFields.Item(8).text = Chr(34) & "Refd. By" & Chr(34)
            '--------------------------------------------------------------------
            Report46.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report46.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report46.FormulaFields.Item(11).text = Chr(34) & "1. Organism Isolated" & Chr(34)
            Report46.FormulaFields.Item(12).text = Chr(34) & "2. Sensitivity Test" & Chr(34)
            Report46.FormulaFields.Item(13).text = Chr(34) & "Checked By" & Chr(34)

            Call Flush_Doc_Name
            Report46.Text2.SetText StDoc_Name

            Report46.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report46.Database.SetDataSource rs
            CRViewer1.ReportSource = Report46

            Case 47
             StrStdt = Trim(Format(frmPat_Info_VAT.stDt, "yyyy-mm-dd"))
             StrSttime = Trim(Format(frmPat_Info_VAT.stDT_TM, "hh:mm"))
             StDate_TM = StrStdt + Space(1) + StrSttime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++

            '++++++for Ending Date and Time++++++++++++++

             StrEddt = Trim(Format(frmPat_Info_VAT.edDt, "yyyy-mm-dd"))
             StrEdtime = Trim(Format(frmPat_Info_VAT.edDT_TM, "hh:mm"))
             EdDate_TM = StrEddt + Space(1) + StrEdtime
            '++++++++++end+++++++++++++++++++++++++++++++++++++++



            '///////for use CONVERT FUNCTION////////////////////
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            '////////////////////////////////////////////////

            Report30.FormulaFields.Item(1).text = Chr(34) & Format(StDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
            Report30.FormulaFields.Item(2).text = Chr(34) & Format(EdDate_TM, "dd/mm/yyyy hh:mm AMPM") & Chr(34)
            Report30.FormulaFields.Item(5).text = Chr(34) & Format(rDaily_Statement.stDt, "YYYY") & Chr(34)

            Report30.DiscardSavedData

            rs.Open "exec Daily_Stat_vat '" & StDate_TM & "','" & EdDate_TM & "'", strcn.Connection
            Report30.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report30

      Case 48
            con.connectionstring = strcn.Connection
            con.Open
            Set cmd.ActiveConnection = con

            Report47.DiscardSavedData

            rs.Open "exec st_balance 1", strcn.Connection

            Report47.Database.SetDataSource rs
            con.Close
            CRViewer1.ReportSource = Report47

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




