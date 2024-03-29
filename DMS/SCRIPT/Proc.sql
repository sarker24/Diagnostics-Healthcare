if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Advance_Coll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Advance_Coll]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Advance_Coll1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Advance_Coll1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CR_Date]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[CR_Date]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Commission_Per_Select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Commission_Per_Select]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Commission_Select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Commission_Select]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Commission_Select3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Commission_Select3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Commission_Select4]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Commission_Select4]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Commission_Select5]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Commission_Select5]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Commission_Select6]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Commission_Select6]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Daily_Stat]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Daily_Stat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Daily_Stat1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Daily_Stat1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Daily_Stat2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Daily_Stat2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Daily_Stat_Test]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Daily_Stat_Test]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Daily_Stat_VAT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Daily_Stat_VAT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Del_Doc_New]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Del_Doc_New]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Del_Report]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Del_Report]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Delete_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Delete_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Doc_SELECT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Doc_SELECT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Doc_name]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Doc_name]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Due_Coll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Due_Coll]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Due_Coll_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Due_Coll_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Due_Doc_Pat]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Due_Doc_Pat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Due_Pat]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Due_Pat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FLUSH_Com]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[FLUSH_Com]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Flush_Test_Result]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Flush_Test_Result]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Font_IU]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Font_IU]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetTestName]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[GetTestName]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Ins_Into_VAT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Ins_Into_VAT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Inv_NO_SELECT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Inv_NO_SELECT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Item_Info_IUD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Item_Info_IUD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Leave_As_Cash_IUD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Leave_As_Cash_IUD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Leave_Balance]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Leave_Balance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Leave_Balance1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Leave_Balance1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Make_Pat_ID1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Make_Pat_ID1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Make_Pat_ID_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Make_Pat_ID_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[New_Doc_Select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[New_Doc_Select]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PAT_INFO_MAIN_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PAT_INFO_MAIN_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pat_Info_SELECT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pat_Info_SELECT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pat_Info_SELECT1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pat_Info_SELECT1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pat_Info_SELECT_VAT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pat_Info_SELECT_VAT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pat_Info_Sub1_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pat_Info_Sub1_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pat_Info_Sub1_Delete1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pat_Info_Sub1_Delete1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pat_Type]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pat_Type]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_Auto]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_Auto]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_Current_Stock]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_Current_Stock]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_FLUSH]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_FLUSH]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_FLUSH1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_FLUSH1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_FLUSH_TN]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_FLUSH_TN]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_FLUSH_VAT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_FLUSH_VAT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_Pur_Iss_Contrast]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_Pur_Iss_Contrast]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_Stock_det]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_Stock_det]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_TOTrate_commission]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_TOTrate_commission]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_comm_main_FLUSH]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_comm_main_FLUSH]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_commission_flush]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_commission_flush]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_flush_unique_id]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_flush_unique_id]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_vat_make_serial]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_vat_make_serial]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Report_All_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Report_All_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Report_All_Delete1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Report_All_Delete1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Report_All_Delete2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Report_All_Delete2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Report_All_SELECT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Report_All_SELECT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Report_All_SELECT3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Report_All_SELECT3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Report_All_Select1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Report_All_Select1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Report_All_Select2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Report_All_Select2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Report_All_Select4]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Report_All_Select4]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpr_Doc_Pay3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpr_Doc_Pay3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Booth]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Booth]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Booth1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Booth1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Cancer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Cancer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Doc_Pay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Doc_Pay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Doc_Pay1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Doc_Pay1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Doc_Pay_new]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Doc_Pay_new]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Doc_pay2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Doc_pay2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Doctor_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Doctor_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Doctor_Info_New]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Doctor_Info_New]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Pat_info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Pat_info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Test_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Test_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_VAT_Pat]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_VAT_Pat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_for_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_for_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SELECT_Leave]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SELECT_Leave]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[S_Name_Select1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[S_Name_Select1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[S_name_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[S_name_select]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Search_Leave_Type]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Search_Leave_Type]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Search_Leave_as_cash]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Search_Leave_as_cash]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Search_Pat_ID]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Search_Pat_ID]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Search_Pat_ID1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Search_Pat_ID1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Search_Pat_ID2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Search_Pat_ID2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Search_Pat_Type]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Search_Pat_Type]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Select_New_Doc_Name]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Select_New_Doc_Name]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Select_Paid]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Select_Paid]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Select_Soft_Sucurity]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Select_Soft_Sucurity]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Select_Soft_Sucurity1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Select_Soft_Sucurity1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Select_VAT_Pat]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Select_VAT_Pat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[St_Balance]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[St_Balance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Stock_In_IUD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Stock_In_IUD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Stock_Out_IUD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Stock_Out_IUD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Stock_Status]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Stock_Status]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sup_Info_IUD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Sup_Info_IUD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select10]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select10]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select11]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select11]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select12]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select12]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select13]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select13]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select15]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select15]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select16]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select16]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select17]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select17]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select18]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select18]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select19]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select19]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Test_Result_Select8]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Test_Result_Select8]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[U_PAT_Disc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[U_PAT_Disc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[U_PAT_Pay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[U_PAT_Pay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[U_PAT_Test_Code]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[U_PAT_Test_Code]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[U_VAT_ID]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[U_VAT_ID]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Vat_Setup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Vat_Setup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Vat_Setup_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Vat_Setup_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[commission_pay_flush]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[commission_pay_flush]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[commission_select2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[commission_select2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[doc_comm_pay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[doc_comm_pay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[doc_comm_pay1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[doc_comm_pay1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[doc_comm_pay2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[doc_comm_pay2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[doc_comm_payU]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[doc_comm_payU]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[m_name_select]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[m_name_select]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pass_para]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pass_para]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_COMPANY_INFO]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_COMPANY_INFO]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Commission_Details]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Commission_Details]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Commission_Pay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Commission_Pay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Commission_Per]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Commission_Per]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_DOCTOR_INFO]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_DOCTOR_INFO]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_DOCTOR_INFO_NEW]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_DOCTOR_INFO_NEW]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_DOCTOR_INFO_NEW1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_DOCTOR_INFO_NEW1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_DOCTOR_INFO_NEW2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_DOCTOR_INFO_NEW2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_DOCTOR_INFO_NEW3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_DOCTOR_INFO_NEW3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Emp_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Emp_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Leave_IUD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Leave_IUD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Leave_setup_IUD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Leave_setup_IUD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_PAT_INFO_MAIN]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_PAT_INFO_MAIN]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_PAT_INFO_MAIN_UD]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_PAT_INFO_MAIN_UD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_PAT_INFO_SUB1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_PAT_INFO_SUB1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_PAT_INFO_SUB2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_PAT_INFO_SUB2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_PAT_INFO_SUB3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_PAT_INFO_SUB3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Report_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Report_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Soft_Security]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Soft_Security]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Soft_Security1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Soft_Security1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_TEST_INFO_MAIN]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_TEST_INFO_MAIN]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_TEST_INFO_RATE]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_TEST_INFO_RATE]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_TEST_INFO_SUB]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_TEST_INFO_SUB]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_TEST_RESULT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_TEST_RESULT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_TEST_RESULT1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_TEST_RESULT1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_Test_Info_FLUSH]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_Test_Info_FLUSH]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_USR_INFO]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_USR_INFO]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_commpay__select1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_commpay__select1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_micropass]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_micropass]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_name_SELECT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_name_SELECT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_name_SELECT1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_name_SELECT1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_pass_entry]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_pass_entry]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_security]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_security]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_security_entry]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_security_entry]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rptTest_State]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rptTest_State]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_found]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_found]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_Info_SELECT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_Info_SELECT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_Result_Select7]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_Result_Select7]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_result_SELECT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_result_SELECT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_result_SELECT1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_result_SELECT1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_result_SELECT14]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_result_SELECT14]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_result_SELECT2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_result_SELECT2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_result_SELECT3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_result_SELECT3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_result_SELECT4]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_result_SELECT4]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_result_SELECT5]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_result_SELECT5]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_result_SELECT6]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_result_SELECT6]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[test_result_SELECT9]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[test_result_SELECT9]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


/****** Object:  Stored Procedure dbo.Advance_Coll    Script Date: 03/04/2002 4:10:13 PM ******/
CREATE        PROCEDURE Advance_Coll

@status int,
@u_id varchar(10),
@StDate as datetime,
@EdDate as datetime

AS
set nocount on

if @status=1
begin
--drop table #Adv_Coll1
--drop table #tmp
--drop table #tmp1
--drop table #tmp2
--drop table #tmp3
--drop table #tmp4
--drop table #Adv_Coll

--declare @StDate as datetime
--declare @EdDate as datetime

--set @StDate='2002-03-06 11:30:00.000'
--set @EdDate='2002-03-08 11:34:00.000'


CREATE TABLE [#Adv_Coll1] (
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL ,
	[usr_name] [varchar] (1000) NULL
) ON [PRIMARY]


-----adv----------------------------------
select c.pat_id,adv=isnull(sum(c.adv),0),
doc_name=isnull((select doc_name from doctor_info e where
e.refer_code=a.refer_code),(select doc_name from doctor_info_new f where
f.pat_id=a.pat_id))
,m.u_name
into #tmp2
from pat_info_sub2 c,pat_info_main a,micropass m
where m.u_id=c.uid and c.uid=@u_id 
and a.pat_id=c.pat_id and c.adv<>0
and c.type='adv'
and c.dt1 between @StDate and @EdDate
group by c.pat_id,m.u_name,a.refer_code,a.pat_id

--select*from #tmp2

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,0,0,0,0,adv,0,doc_name,u_name from #tmp2

--select*from #Adv_Coll1

--<<<<<<<Advance<<<<<<<<<<<<<<---------
-->>>>--due-->>>>>>>>>>>>>>>>>>>>>>---------------------------------
select c.pat_id,due=isnull(sum(c.adv),0),t.doc_name,t.u_name
into #tmp3
from pat_info_sub2 c,#tmp2 t
where t.pat_id=c.pat_id and c.type='due'
group by c.pat_id,c.type,t.doc_name,t.u_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,0,0,0,0,0,adv=isnull(sum(due),0),doc_name,u_name from #tmp3 group by pat_id,doc_name,u_name

--select*from #tmp3
--SELECT*FROM #Adv_Coll1
---<<<<<-----due--------------------------------------------------

--->>>---collect_fee->>>>>>>>>>-----------------------
select c.pat_id,collect_fee=isnull(sum(c.collect_fee),0),t.doc_name,t.u_name
into #tmp4
from pat_info_sub2 c,#tmp2 t
where c.pat_id=t.pat_id
group by c.pat_id,t.doc_name,t.u_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,0,0,0,collect_fee,0,0,doc_name,u_name from #tmp4 
--group by pat_id,doc_name,u_name
--SELECT*FROM #Adv_Coll1
--SELECT*FROM #tmp4
--SELECT*FROM #tmp2
--select*from pat_info_sub2
------<<<<<<collection<<<<<<<<<------------------------------

----test rate>>>>>>>>>>----
select b.pat_id,test_rate=isnull(sum(b.test_rate),0),t.doc_name,t.u_name into #tmp 
from pat_info_sub1 b,#tmp2 t
where b.pat_id=t.pat_id
group by b.pat_id,t.doc_name,t.u_name
--select*from #tmp

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,test_rate,0,0,0,0,0,doc_name,u_name from #tmp
--SELECT*FROM #Adv_Coll1
---<<<<<<<<<<<<<----

--->>>>discount--
select d.pat_id,disc=sum(d.disc)
into #tttmp1
from pat_info_sub3 d,#tmp2 t
where d.pat_id=t.pat_id
group by d.pat_id
--<<-end

---VAT,disc>>>>>>>>>>>>>>>>>>

select distinct a.pat_id,
a.vat_amt,d.disc,t.doc_name,t.u_name
into #tmp1
from pat_info_main a,#tttmp1 d,
pat_info_sub2 c,#tmp2 t
where a.pat_id=d.pat_id and a.pat_id=c.pat_id 
and a.pat_id=t.pat_id
--group by a.pat_id,
--a.vat_amt,
--d.disc,
--t.doc_name,t.u_name
--select*from #tmp1

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,0,vat_amt,disc,0,0,0,doc_name,u_name from #tmp1


---<<<<<<<<<<<<<<<<<------------
---->>Create Auto_No---->>>>>>>>---------------
CREATE TABLE [#Adv_Coll] (
	[sl_no] [int] IDENTITY (1, 1) NOT NULL ,
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL, 
	[usr_name] [varchar] (1000) NULL

) ON [PRIMARY]

insert into #adv_coll(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,test_rate=sum(test_rate),VAT_amt=sum(VAT_amt),
disc=sum(disc),coll_fee=sum(coll_fee),adv_coll=sum(adv_coll),
due_coll=sum(due_coll),doc_name,usr_name from #Adv_Coll1
group by pat_id,doc_name,usr_name
---<<<<<<<<<<<<<<<<
-->>>>>>final>>>>>>>>>>
--select sl_no,pat_id,test_rate,VAT_amt,total=(test_rate+VAT_amt+coll_fee-disc),due_amt=(test_rate+VAT_amt+coll_fee-disc-adv_coll-due_coll),disc,coll_fee,adv_coll,due_coll,doc_name,usr_name from #adv_coll

select t.sl_no,t.pat_id,t.test_rate,t.VAT_amt,
total=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc),
due_amt=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc-t.adv_coll-t.due_coll),
t.disc,t.coll_fee,t.adv_coll,t.due_coll,t.doc_name,t.usr_name,a.pat_id1
into #Adv_Coll2 from #adv_coll t,pat_info_main a where t.pat_id=a.pat_id

update #Adv_Coll2 set pat_id1=pat_id where pat_id1=''

select*from #Adv_Coll2


end


set nocount off



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






CREATE       PROCEDURE Advance_Coll1

@status int,
@StDate as datetime,
@EdDate as datetime

AS

set nocount on

if @status=1
begin


CREATE TABLE [#Adv_Coll1] (
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL ,
	[usr_name] [varchar] (1000) NULL
) ON [PRIMARY]


-----adv----------------------------------
select c.pat_id,adv=isnull(sum(c.adv),0),
doc_name=isnull((select doc_name from doctor_info e where
e.refer_code=a.refer_code),(select doc_name from doctor_info_new f where
f.pat_id=a.pat_id))
,m.u_name
into #tmp2
from pat_info_sub2 c,pat_info_main a,micropass m
where m.u_id=c.uid
--and c.uid=@u_id 
and a.pat_id=c.pat_id and c.adv<>0
and c.type='adv'
and c.dt1 between @StDate and @EdDate
group by c.pat_id,m.u_name,a.refer_code,a.pat_id

--select*from #tmp2
--select*from pat_info_sub2

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,0,0,0,0,adv,0,doc_name,u_name from #tmp2

--select*from #Adv_Coll1
--<<<<<<<Advance<<<<<<<<<<<<<<---------

-->>>>--due-->>>>>>>>>>>>>>>>>>>>>>---------------------------------
select c.pat_id,due=isnull(sum(c.adv),0),t.doc_name,t.u_name
into #tmp3
from pat_info_sub2 c,#tmp2 t
where t.pat_id=c.pat_id and c.type='due'
group by c.pat_id,c.type,t.doc_name,t.u_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,0,0,0,0,0,adv=isnull(sum(due),0),doc_name,u_name from #tmp3 group by pat_id,doc_name,u_name

--select*from #tmp3
--SELECT*FROM #Adv_Coll1
--select*from pat_info_sub2
---<<<<<-----due--------------------------------------------------

--->>>---collect_fee->>>>>>>>>>-----------------------
select c.pat_id,collect_fee=isnull(sum(c.collect_fee),0),t.doc_name,t.u_name
into #tmp4
from pat_info_sub2 c,#tmp2 t
where c.pat_id=t.pat_id
group by c.pat_id,t.doc_name,t.u_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,0,0,0,collect_fee,0,0,doc_name,u_name from #tmp4 
--group by pat_id,doc_name,u_name
--SELECT*FROM #Adv_Coll1
--SELECT*FROM #tmp4
--SELECT*FROM #tmp2
--select*from pat_info_sub2
------<<<<<<collection<<<<<<<<<------------------------------


----test rate>>>>>>>>>>----
select b.pat_id,test_rate=isnull(sum(b.test_rate),0),t.doc_name,
t.u_name into #tmp 
from pat_info_sub1 b,#tmp2 t
where b.pat_id=t.pat_id
group by b.pat_id,t.doc_name,t.u_name
--select*from #tmp

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,test_rate,0,0,0,0,0,doc_name,u_name from #tmp
--SELECT*FROM #Adv_Coll1
--SELECT*FROM #tmp
---<<<<<<<<<<<<<----

-->>--only for Discount--
select d.pat_id,disc=sum(d.disc)
into #tmpppp
from pat_info_sub3 d,#tmp2 t
where t.pat_id=d.pat_id
group by d.pat_id

--end--
--->>VAT,disc>>>>>>>>>>>>>>>>>>
select a.pat_id,
a.vat_amt,d.disc,t.doc_name,t.u_name
into #tmp1
from pat_info_main a,#tmpppp d,
pat_info_sub2 c,#tmp2 t
where a.pat_id=d.pat_id and a.pat_id=c.pat_id 
and a.pat_id=t.pat_id
group by a.pat_id,
a.vat_amt,d.disc,
t.doc_name,t.u_name
--select*from #tmp1

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,0,vat_amt,disc,0,0,0,doc_name,u_name from #tmp1


---<<<<<<<<<<<<<<<<<------------
---->>Create Auto_No---->>>>>>>>---------------
CREATE TABLE [#Adv_Coll] (
	[sl_no] [int] IDENTITY (1, 1) NOT NULL ,
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL, 
	[usr_name] [varchar] (1000) NULL

) ON [PRIMARY]
 insert into #adv_coll(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,usr_name)
select pat_id,test_rate=sum(test_rate),VAT_amt=sum(VAT_amt), disc=sum(disc),coll_fee=sum(coll_fee),adv_coll=sum(adv_coll),
due_coll=sum(due_coll),doc_name,usr_name from #Adv_Coll1
group by pat_id,doc_name,usr_name
---<<<<<<<<<<<<<<<<
-->>>>>>final>>>>>>>>>>
--select sl_no,pat_id,test_rate,VAT_amt,total=(test_rate+VAT_amt+coll_fee-disc),due_amt=(test_rate+VAT_amt+coll_fee-disc-adv_coll-due_coll),disc,coll_fee,adv_coll,due_coll,doc_name,usr_name from #adv_coll

select t.sl_no,t.pat_id,t.test_rate,t.VAT_amt,
total=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc),
due_amt=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc-t.adv_coll-t.due_coll),
t.disc,t.coll_fee,t.adv_coll,t.due_coll,t.doc_name,t.usr_name,a.pat_id1
into #Adv_Coll2 from #adv_coll t,pat_info_main a where t.pat_id=a.pat_id

update #Adv_Coll2 set pat_id1=pat_id where pat_id1=''
select*from #Adv_Coll2


end

set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROC CR_Date
As

select getdate() CrDATE


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE Commission_Per_Select

@status int,
@type varchar(10),
@refer_code varchar (10)

AS

if @status=1
begin
	select*from commission_per where type=@type and refer_code=@refer_code
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE PROCEDURE Commission_Select


@refer_code varchar(10)

 AS

select a.refer_code,
(select b.doc_name from doctor_info b where a.refer_code=b.refer_code) as doc_name,"New" from
pat_info_main a where a.refer_code=@refer_code

union

select c.refer_code,
(select d.doc_name from doctor_info d where c.refer_code=d.refer_code) as doc_name,"Old" from
commission_details c where c.refer_code=@refer_code








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Commission_Select3    Script Date: 18/10/2001 6:18:21 PM ******/

CREATE PROCEDURE Commission_Select3

@status int,
@refer_code varchar(10),
--@m_code varchar(2),
@s_code varchar(3)

AS
if @status=1

begin


  /*      select refer_code as Doctor_ID,pat_id as Patient_ID,s_code as Test_code,
        commission as Amount,dt as Entry_Date,note as Note,
       Test_name=(select s_name from test_info_sub where commission_details.s_code=test_info_sub.s_code),
       Test_rate=(select test_rate from pat_info_sub1 where s_code=@s_code
       and refer_code=@refer_code
       and pat_info_sub1.pat_id=commission_details.pat_id)
       from commission_details where refer_code=@refer_code and s_code=@s_code
       and pat_id not in(select pat_id 
       from commission_pay where commission_pay.pat_id=commission_details.pat_id and refer_code=@refer_code)
*/

select refer_code as Doctor_ID,pat_id as Patient_ID,m_code as Main,s_code as Test_code,
commission as Amount,dt as Entry_Date,note as Note,
Test_name=(select s_name from test_info_sub where 
commission_details.m_code=test_info_sub.m_code and
commission_details.s_code=test_info_sub.s_code),
Test_rate=(select test_rate from pat_info_sub1 where 
commission_details.m_code=pat_info_sub1.m_code and
commission_details.s_code=pat_info_sub1.s_code
and pat_info_sub1.pat_id=commission_details.pat_id)
from commission_details where refer_code=@refer_code and s_code=@s_code
and pat_id not in(select pat_id 
from commission_pay where commission_pay.pat_id=commission_details.pat_id 
and commission_pay.m_code=commission_details.m_code
and commission_pay.s_code=commission_details.s_code
and refer_code=@refer_code)

end










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Commission_Select4    Script Date: 15/10/2001 6:26:43 PM ******/
CREATE PROCEDURE Commission_Select4

@status int,
@refer_code varchar(10),
@pat_id varchar(10),
@s_code varchar(3)

AS

if @status=1
begin
	select*from commission_details where refer_code=@refer_code and pat_id=@pat_id and s_code=@s_code
end

if @status=2
begin
	delete from commission_details where refer_code=@refer_code and pat_id=@pat_id and s_code=@s_code
end
if @status=3
begin
	select*from commission_pay where refer_code=@refer_code and pat_id=@pat_id and s_code=@s_code
end
if @status=4
begin
	delete from commission_pay where refer_code=@refer_code and pat_id=@pat_id and s_code=@s_code
end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE PROCEDURE Commission_Select5

@status int,
@refer_code varchar(10),
@s_code varchar(3)

AS
if @status=1

begin
	select refer_code as Doctor_ID,pat_id as Patient_ID,m_code as Main,s_code as Test_code,
commission as Amount,dt as Pay_Date,note,
Test_rate=(select test_rate from pat_info_sub1 
where commission_details.m_code=pat_info_sub1.m_code and s_code=@s_code
             and refer_code=@refer_code
             and pat_info_sub1.pat_id=commission_details.pat_id) from commission_details
             where refer_code=@refer_code and s_code=@s_code

end

if @status=2

begin

--        select a.refer_code,a.pat_id,a.m_code,a.s_code,a.paid,a.dt,a.cleared,a.note,
--        Test_rate=(select top 1 test_rate from pat_info_sub1 where s_code=@s_code and refer_code=@refer_code)
--        from commission_pay a where a.refer_code=@refer_code and a.s_code=@s_code

select a.refer_code,a.pat_id,a.m_code,a.s_code,a.paid,a.dt,a.cleared,a.note,
Test_rate=(select top 1 test_rate from pat_info_sub1 where 
a.s_code=pat_info_sub1.s_code
and a.m_code=pat_info_sub1.m_code
and a.pat_id=pat_info_sub1.pat_id)
from commission_pay a where a.refer_code=@refer_code and a.s_code=@s_code

end























GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Commission_Select6    Script Date: 18/10/2001 6:18:22 PM ******/
CREATE PROCEDURE Commission_Select6

@status int,
@refer_code varchar(10),
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3)

AS


if @status=1
begin
	delete from commission_details where refer_code=@refer_code and pat_id=@pat_id and m_code=@m_code and s_code=@s_code
end
if @status=2
begin
	delete from commission_pay where refer_code=@refer_code and pat_id=@pat_id and m_code=@m_code and s_code=@s_code
end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




/****** Object:  Stored Procedure dbo.Daily_Stat    Script Date: 03/04/2002 4:10:13 PM ******/
CREATE  PROCEDURE Daily_Stat 

@StDateTM datetime,
@EdDateTM datetime

AS
set nocount on

-------from pat_info_sub1 test rate------------
--drop table #tmp
select pat_id,s_test_rate=sum(test_rate),tmp_dt 
into #tmp
from pat_info_sub1
where tmp_dt between @StDateTM and @EdDateTM
group by pat_id,tmp_dt

delete from daily_state
insert into daily_state(dt,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll)
select tmp_dt,sum(s_test_rate),0,0,0,0,0  from #tmp group by tmp_dt

-----------------------------------------------------------------

--drop table #tmp1

select tmp_dt,vat_amt 
into #tmp1 from pat_info_main
where tmp_dt between @StDateTM and @EdDateTM
group by tmp_dt,vat_amt

insert into daily_state(dt,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll)
select tmp_dt,0,sum(vat_amt),0,0,0,0 from #tmp1 group by tmp_dt

----==============from pat_info_sub3===================================================
--drop table #tmp2

select tmp_dt,disc 
into #tmp2 from pat_info_sub3
where tmp_dt between @StDateTM and @EdDateTM
group by tmp_dt,disc

insert into daily_state(dt,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll)
select tmp_dt,0,0,sum(disc),0,0,0 from #tmp2 group by tmp_dt

---==============from pat_info_sub2 advance collection===========================
--drop table #tmp3
select tmp_dt,
adv into #tmp3
from pat_info_sub2 where type='adv'
and tmp_dt between @StDateTM and @EdDateTM

insert into daily_state(dt,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll)
select tmp_dt,0,0,0,0,sum(adv),0 from #tmp3 group by tmp_dt

--++++++++++++pat_info_sub2 DUE+++++++++++++++++++++++++++

--drop table #tmp4
select tmp_dt,
adv into #tmp4
from pat_info_sub2 where type='due'
and tmp_dt between @StDateTM and @EdDateTM
group by tmp_dt,adv

insert into daily_state(dt,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll)
select tmp_dt,0,0,0,0,0,sum(adv) from #tmp4 group by tmp_dt

--#####--pat_info_sub2 collection fee--------------------
--drop table #tmp5
select tmp_dt,
collect_fee into #tmp5
from pat_info_sub2 where
tmp_dt between @StDateTM and @EdDateTM
group by tmp_dt,collect_fee

insert into daily_state(dt,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll)
select tmp_dt,0,0,0,sum(collect_fee),0,0 from #tmp5 group by tmp_dt

--select * from daily_state 

set nocount off








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO



/****** Object:  Stored Procedure dbo.Daily_Stat1    Script Date: 03/04/2002 4:10:12 PM ******/
CREATE PROCEDURE Daily_Stat1


AS

select * from daily_state
--dt,test_rate,VAT_amt,disc,coll_fee,
--adv_coll,due_coll
--,uid,
--u_name=(select u_name from micropass where micropass.u_id=daily_state.uid) 
--from daily_state







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO







/****** Object:  Stored Procedure dbo.Daily_Stat2    Script Date: 03/04/2002 4:10:13 PM ******/
CREATE    PROCEDURE Daily_Stat2

@StDateTM datetime,
@EdDateTM datetime

AS
set nocount on


CREATE TABLE [#daily_state] (
	[dt] [datetime] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[collect_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
) ON [PRIMARY]



-------from pat_info_sub1 test rate------------

select pat_id,s_test_rate=sum(test_rate),tmp_dt 
into #tmp
from pat_info_sub1
where dt1 between @StDateTM and @EdDateTM
group by pat_id,tmp_dt


insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,sum(s_test_rate),0,0,0,0,0  from #tmp group by tmp_dt

-----------------------------------------------------------------

select tmp_dt,vat_amt 
into #tmp1 from pat_info_main
where dt1 between @StDateTM and @EdDateTM
--group by tmp_dt,vat_amt

insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,0,sum(vat_amt),0,0,0,0 from #tmp1 group by tmp_dt

----==============from pat_info_sub3===================================================

select tmp_dt,disc 
into #tmp2 from pat_info_sub3
where dt1 between @StDateTM and @EdDateTM
--group by tmp_dt,disc

insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,0,0,sum(disc),0,0,0 from #tmp2 group by tmp_dt

---==============from pat_info_sub2 advance collection===========================
select tmp_dt,
adv into #tmp3
from pat_info_sub2 where type='adv'
and dt1 between @StDateTM and @EdDateTM

insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,0,0,0,0,sum(adv),0 from #tmp3 group by tmp_dt

--++++++++++++pat_info_sub2 DUE+++++++++++++++++++++++++++

select dt2,
adv into #tmp4
from pat_info_sub2 where type='due'
---and dt1 between @StDateTM and @EdDateTM
and dt between @StDateTM and @EdDateTM
--group by tmp_dt,adv

insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select dt2,0,0,0,0,0,sum(adv) from #tmp4 group by dt2

--#####--pat_info_sub2 collection fee--------------------
/*
select tmp_dt,
collect_fee into #tmp5
from pat_info_sub2 where
dt1 between @StDateTM and @EdDateTM
group by tmp_dt,collect_fee
*/
select a.tmp_dt,collect_fee=sum(a.collect_fee) into #tmp5 from pat_info_sub2 a,#tmp t where a.pat_id=t.pat_id group by a.tmp_dt

insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,0,0,0,sum(collect_fee),0,0 from #tmp5 group by tmp_dt

----**end PART-1***
---*********PART-2****

CREATE TABLE [#tmp6] (
	[id_seed] [int] IDENTITY (1, 1) NOT NULL ,
	[dt] [datetime] NULL,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[collect_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[total_coll] [money] NULL ,
	[grand_tot] [money] NULL ,
	[average] [money] NULL 
)

insert into #tmp6
SELECT dt,test_rate=SUM(test_rate),
VAT_amt=SUM(VAT_amt),disc=SUM(disc),collect_fee=SUM(collect_fee),adv_coll=SUM(adv_coll),
due_coll=SUM(due_coll),
total_coll=(SUM(adv_coll)+SUM(due_coll)),0,0
from #daily_state
group by dt

declare @c as char(1)
declare @id_seed  as int
declare @total_coll as money
declare @average  as money

declare @gt as money
declare @seed  as int
set @seed=1
set @gt=0

DECLARE Csr CURSOR FOR
select id_seed,total_coll,average from #tmp6

OPEN Csr
set @c='1'
WHILE @c = '1'
begin   
	FETCH NEXT FROM Csr into @id_seed,@total_coll,@average
	if @@FETCH_STATUS = 0 
    	begin
		set @gt=@gt+@total_coll
		if @id_seed=1
		begin	
			update #tmp6 set grand_tot=@total_coll,average=@total_coll where id_seed=1
		end
	
		else
		begin

			update #tmp6 set grand_tot=@gt,average=@gt/@seed where id_seed=@seed
		end

		set @seed=@seed+1

	end
	else set @c = '0'
end
CLOSE Csr
DEALLOCATE Csr


select id_seed,dt,test_rate,Tot_Bill=(test_rate+VAT_amt+collect_fee),VAT_amt,disc,collect_fee,adv_coll,due_coll,total_coll,grand_tot,average from #tmp6



set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Daily_Stat_Test    Script Date: 15/04/2002 11:52:56 AM ******/

CREATE  PROCEDURE Daily_Stat_Test

@State int,
@m_name varchar (30),
@StDateTM datetime,
@EdDateTM datetime

AS
set nocount on

if @State=1
begin
/*
	select a.m_code,t.m_name,a.s_code,b.s_name,a.test_rate 
	from pat_info_sub1 a,test_info_main t,test_info_sub b
	where t.m_code=a.m_code
	and b.m_code=a.m_code
	and a.s_code=b.s_code
	and a.dt1 between @StDateTM and @EdDateTM
*/

--	declare @m_name varchar (30)
--declare @StDateTM datetime
--declare @EdDateTM datetime
--set @m_name='BIOCHEMICAL EXAMINATION'
--set @StDateTM='2002-03-16 00:12:45.540'
--set @EdDateTM='2002-05-16 00:00:45.540'
---------part 1-----------
--drop table #pat_info_sub1
select a.m_code,a.s_code,count_SNAME=count(a.s_code) into #pat_info_sub1
from pat_info_sub1 a
where a.dt1 between @StDateTM and @EdDateTM
group by a.m_code,a.s_code,a.test_rate
----end part 1---
--part 2----
select f.m_code,f.s_code,t.m_name,b.s_name,f.count_SNAME
from test_info_main t,#pat_info_sub1 f,test_info_sub b
where f.m_code=t.m_code and b.m_code=t.m_code
and f.s_code=b.s_code
---end part 2--






end

if @State=2
begin
/*	
select a.m_code,t.m_name,a.s_code,b.s_name,a.test_rate 
	from pat_info_sub1 a,test_info_main t,test_info_sub b
	where t.m_code=a.m_code
	and b.m_code=a.m_code
	and a.s_code=b.s_code
	and a.dt1 between @StDateTM and @EdDateTM
	and t.m_name=@m_name
*/
--	drop table #pat_info_sub1
select a.m_code,a.s_code,count_SNAME=count(a.s_code) into #pat_info_s1
from pat_info_sub1 a,test_info_main m
where m.m_name=@m_name and a.m_code=m.m_code and a.dt1 between @StDateTM and @EdDateTM
group by a.m_code,a.s_code
--select*from #pat_info_sub1


----end part 1---
--part 2----
select f.m_code,f.s_code,t.m_name,b.s_name,f.count_SNAME
from test_info_main t,#pat_info_s1 f,test_info_sub b
where f.m_code=t.m_code and b.m_code=t.m_code
and f.s_code=b.s_code
---end part 2--




end

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE     PROCEDURE Daily_Stat_VAT

@StDateTM datetime,
@EdDateTM datetime

AS
set nocount on


CREATE TABLE [#daily_state] (
	[dt] [datetime] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[collect_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
) ON [PRIMARY]



-------from pat_info_sub1 test rate------------

select pat_id,s_test_rate=sum(test_rate),tmp_dt 
into #tmp
from pat_info_sub1_VAT
where dt1 between @StDateTM and @EdDateTM
group by pat_id,tmp_dt


insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,sum(s_test_rate),0,0,0,0,0  from #tmp group by tmp_dt

-----------------------------------------------------------------

select tmp_dt,vat_amt 
into #tmp1 from pat_info_main_VAT
where dt1 between @StDateTM and @EdDateTM
--group by tmp_dt,vat_amt

insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,0,sum(vat_amt),0,0,0,0 from #tmp1 group by tmp_dt

-----==============from pat_info_sub3===================================================

select tmp_dt,disc 
into #tmp2 from pat_info_sub3_VAT
where dt1 between @StDateTM and @EdDateTM
--group by tmp_dt,disc

insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,0,0,sum(disc),0,0,0 from #tmp2 group by tmp_dt

---==============from pat_info_sub2 advance collection===========================
select tmp_dt,
adv into #tmp3
from pat_info_sub2_VAT where type='adv'
and dt1 between @StDateTM and @EdDateTM

insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,0,0,0,0,sum(adv),0 from #tmp3 group by tmp_dt

--++++++++++++pat_info_sub2 DUE+++++++++++++++++++++++++++

select dt2,
adv into #tmp4
from pat_info_sub2_VAT where type='due'
and dt between @StDateTM and @EdDateTM
--select*from pat_info_sub2_VAT
insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select dt2,0,0,0,0,0,sum(adv) from #tmp4 group by dt2

--#####--pat_info_sub2 collection fee--------------------

select a.tmp_dt,collect_fee=sum(a.collect_fee) into #tmp5 from pat_info_sub2_VAT a,#tmp t where a.pat_id=t.pat_id group by a.tmp_dt

insert into #daily_state(dt,test_rate,VAT_amt,disc,collect_fee,adv_coll,due_coll)
select tmp_dt,0,0,0,sum(collect_fee),0,0 from #tmp5 group by tmp_dt

----**end PART-1***
---*********PART-2****

CREATE TABLE [#tmp6] (
	[id_seed] [int] IDENTITY (1, 1) NOT NULL ,
	[dt] [datetime] NULL,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[collect_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[total_coll] [money] NULL ,
	[grand_tot] [money] NULL ,
	[average] [money] NULL 
)

insert into #tmp6
SELECT dt,test_rate=SUM(test_rate),
VAT_amt=SUM(VAT_amt),disc=SUM(disc),collect_fee=SUM(collect_fee),adv_coll=SUM(adv_coll),
due_coll=SUM(due_coll),
total_coll=(SUM(adv_coll)+SUM(due_coll)),0,0
from #daily_state
group by dt

declare @c as char(1)
declare @id_seed  as int
declare @total_coll as money
declare @average  as money

declare @gt as money
declare @seed  as int
set @seed=1
set @gt=0

DECLARE Csr CURSOR FOR
select id_seed,total_coll,average from #tmp6

OPEN Csr
set @c='1'
WHILE @c = '1'
begin   
	FETCH NEXT FROM Csr into @id_seed,@total_coll,@average
	if @@FETCH_STATUS = 0 
    	begin
		set @gt=@gt+@total_coll
		if @id_seed=1
		begin	
			update #tmp6 set grand_tot=@total_coll,average=@total_coll where id_seed=1
		end
	
		else
		begin

			update #tmp6 set grand_tot=@gt,average=@gt/@seed where id_seed=@seed
		end

		set @seed=@seed+1

	end
	else set @c = '0'
end
CLOSE Csr
DEALLOCATE Csr


select id_seed,dt,test_rate,Tot_Bill=(test_rate+VAT_amt+collect_fee),VAT_amt,disc,collect_fee,adv_coll,due_coll,total_coll,grand_tot,average from #tmp6

set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROC Del_Doc_New
@mode int,
@pat_id varchar(2),
@uid varchar(20)

As
if @mode=1
begin

delete from doctor_info_new where pat_id=@pat_id and uid=@uid

end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  PROC Del_Report
@mode int,
@pat_id varchar(10)
As
if @mode=1
begin 
	delete from report_all where pat_id=@pat_id
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




CREATE  PROCEDURE Delete_All


@status int,
@refer_code int

AS

if @status=1

begin
	delete from doctor_info_new where pat_id=@refer_code
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE Doc_SELECT

@Status int,
@refer_code varchar(100)
AS

if @Status=1
begin
	select * from Doctor_Info where refer_code=@refer_code
end

if @Status=2
begin
	select u_name from micropass where u_id=@refer_code
end

if @Status=3
begin
	select * from doctor_info_new where pat_id=@refer_code
end

if @Status=4
begin
	select * from pat_info_main where pat_id=@refer_code
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  PROCEDURE Doc_name

@status int,
@pat_id varchar(50)

--set @pat_id='33672'
--set @m_code='01'
--set @s_code=''

As

IF (select COUNT(doc_name) from doctor_info where refer_code=(
	select refer_code from pat_info_main where pat_id=@pat_id))=0
BEGIN

select doc_name from doctor_info_new where pat_id=@pat_id

END
ELSE
BEGIN
	select doc_name from doctor_info where refer_code=(
	select refer_code from pat_info_main where pat_id=@pat_id)
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE                      PROCEDURE Due_Coll

@u_id varchar (10),
@StDate datetime,
@EdDate datetime

AS
set nocount on


--drop table #Adv_Coll1
--drop table #tmp3
--drop table #Pre_tmp
--drop table #Pre_due_tmp
--drop table #tmp2
--drop table #tmp4
--drop table #tmp
--drop table #tmpp11
--drop table #tmp1
--drop table #Adv_Coll
--drop table #Adv_Coll2





CREATE TABLE [#Adv_Coll1] (
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[adv_coll1] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL,
	[u_name] [varchar] (1000) NULL ,
	[pre_coll] [money] NULL,
	[pre_due_coll] [money] NULL 

) ON [PRIMARY]

-----+++SET USER NAME+++-->>
declare @u_name varchar(1000)
set @u_name=(select u_name from micropass where u_id=@u_id)
---<<+++ END USER NAME ++<<<<


------due->>>>>>>>>>>>>>>>>>>>---------------------
select c.pat_id,due=isnull(sum(c.adv),0),
doc_name=isnull((select doc_name from doctor_info e where
e.refer_code=a.refer_code) ,(select doc_name from doctor_info_new f where
f.pat_id=a.pat_id))
into #tmp3 from pat_info_sub2 c,pat_info_main a
where c.type='due' and c.uid=@u_id
and c.adv<>0
and a.pat_id=c.pat_id
and c.dt between @StDate and @EdDate
group by c.pat_id,
a.refer_code,a.pat_id
insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,0,0,0,adv=due,doc_name,@u_name,0,0 from #tmp3



---*>*>*>*>*PREVIOUS COLLECTION->>>>>>>>>>

select t.pat_id,pre_coll=sum(c.adv),t.doc_name into #Pre_tmp from pat_info_sub2 c,#tmp3 t where
--c.pat_id=t.pat_id and dt<@EdDate group by t.pat_id,t.doc_name
c.pat_id=t.pat_id and dt<@StDate group by t.pat_id,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,0,0,0,0,doc_name,@u_name,pre_coll,0 from #Pre_tmp
---*<*<*<*<*<--END PREVIOUS COLLECTION--<<<<--

---=>=>=>=>=PREVIOUS DUE COLLECTION->>>>>>>>>>

select t.pat_id,pre_due_coll=sum(c.adv),t.doc_name into #Pre_due_tmp from pat_info_sub2 c,#tmp3 t 
where c.pat_id=t.pat_id and c.type='due' 
and dt<@StDate group by t.pat_id,t.doc_name
--and dt<@EdDate group by t.pat_id,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,0,0,0,0,doc_name,@u_name,0,pre_due_coll from #Pre_due_tmp

---=<=<=<=<=<--END PREVIOUS DUE COLLECTION--<<<<--


----ADVANCE collection>>>>>>>>>>>>>---------
select c.pat_id,adv=isnull(sum(c.adv),0),t.doc_name
into #tmp2 from pat_info_sub2 c,#tmp3 t where 
c.type='adv' and  t.pat_id=c.pat_id 
group by c.pat_id,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,0,adv=sum(adv),0,0,doc_name,@u_name,0,0 from #tmp2 group by pat_id,doc_name


--<<<<<<<<--Advance collection---------
---collect_fee->>>>>--------------------
select c.pat_id,collect_fee=isnull(sum(c.collect_fee),0),t.doc_name
into #tmp4
from pat_info_sub2 c,#tmp3 t
where t.pat_id=c.pat_id
group by c.pat_id,c.collect_fee,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,collect_fee,0,0,0,doc_name,@u_name,0,0 from #tmp4

--<<<<<--------------------------------

-----test rate->>>>>>>>--------------------------------------------
select b.pat_id,test_rate=sum(b.test_rate),t.doc_name
into #tmp 
from pat_info_sub1 b,#tmp3 t
where b.pat_id=t.pat_id
group by b.pat_id,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,test_rate,0,0,0,0,0,0,doc_name,@u_name,0,0 from #tmp

-----test rate <<<<<<<<<<<<--------------- 
-->>>-discount----------------------------
---DISCOUNT>>>>>>>>>>>>>-------------
select d.pat_id,disc=sum(d.disc)
into #tmpp11
from pat_info_sub3 d,#tmp3 t
where d.pat_id=t.pat_id
group by d.pat_id
---<<------

---VAT>>>>>>>>>>>>>
select a.pat_id,t.doc_name,a.vat_amt,d.disc
into #tmp1
from pat_info_main a,
#tmpp11 d,pat_info_sub2 c,#tmp3 t
where a.pat_id=t.pat_id
and a.pat_id=d.pat_id and a.pat_id=c.pat_id
group by a.pat_id,a.refer_code,t.doc_name,
a.vat_amt,d.disc

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,vat_amt,disc,0,0,0,0,doc_name,@u_name,0,0 from #tmp1

---<<<<<<<<<<<VAT,DISCOUNT--------


---fo create Auto_No------->>>>>>>>>>>>>>>
CREATE TABLE [#Adv_Coll] (
	[sl_no] [int] IDENTITY (1, 1) NOT NULL ,
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[adv_coll1] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL, 
	[u_name] [varchar] (1000) NULL,
	[pre_coll] [money] NULL,
	[pre_due_coll] [money] NULL

) ON [PRIMARY]

insert into #adv_coll(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,test_rate=sum(test_rate),VAT_amt=sum(VAT_amt),
disc=sum(disc),coll_fee=sum(coll_fee),adv_coll=sum(adv_coll),adv_coll1=sum(adv_coll1),
due_coll=sum(due_coll),doc_name,u_name,pre_coll=SUM(pre_coll),pre_due_coll=sum(pre_due_coll)
from  #Adv_Coll1
group by pat_id,doc_name,u_name
--<<<<<<<<<<---------------------

----final->>>>>>>>>>>>--------------
select t.pat_id,t.test_rate,t.VAT_amt,
total=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc),
t.pre_coll,Due_Amt=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc-t.pre_coll),
--farther_due=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc-t.adv_coll-t.pre_due_coll),
farther_due=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc-t.adv_coll-t.due_coll-t.adv_coll1),
t.disc,t.coll_fee,t.adv_coll,t.due_coll,t.doc_name,t.u_name,a.pat_id1,t.pre_due_coll
into #Adv_Coll2 from #adv_coll t,pat_info_main a 
where t.pat_id=a.pat_id order by t.pat_id

update #Adv_Coll2 set pat_id1=pat_id where pat_id1=''

update #Adv_Coll2 set pre_coll=adv_coll where pre_coll=0 and adv_coll<>0

--select a.pat_id,a.test_rate,a.VAT_amt,a.total,a.pre_coll,Due_Amt=(a.total-a.pre_coll),
--farther_due=(a.farther_due-due),a.disc,a.coll_fee,a.adv_coll,a.due_coll,a.doc_name,
--a.u_name,a.pat_id1,a.pre_due_coll from #Adv_Coll2 a,#tmp31 b
--where a.pat_id=b.pat_id



select pat_id,test_rate,VAT_amt,total,pre_coll,Due_Amt=(total-pre_coll),
farther_due,disc,coll_fee,adv_coll,due_coll,doc_name,
u_name,pat_id1,pre_due_coll,Another_Due=0
into #Adv_Coll3 from #Adv_Coll2

-----*>>>********Search Due who is not this User**************>>>>>>>>>>>
select c.pat_id,due=isnull(sum(c.adv),0),
doc_name=isnull((select doc_name from doctor_info e where
e.refer_code=a.refer_code) ,(select doc_name from doctor_info_new f where
f.pat_id=a.pat_id))
into #tmp31 from pat_info_sub2 c,#tmp3 t,pat_info_main a
where c.type='due' and c.uid<>@u_id
and c.adv<>0 and a.pat_id=c.pat_id and t.pat_id=c.pat_id
and c.dt between @StDate and @EdDate
group by c.pat_id,a.refer_code,a.pat_id

---<<<<<<<<<<<<<<********************************

update #Adv_Coll3
set Another_Due=(select isnull(sum(due),0) from #tmp31)

select pat_id,test_rate,VAT_amt,total,pre_coll,Due_Amt=(total-pre_coll),
farther_due=(farther_due-Another_Due),disc,coll_fee,adv_coll,due_coll,doc_name,
u_name,pat_id1,pre_due_coll
from #Adv_Coll3

-----<<<*************end****************



set nocount off







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE            PROCEDURE Due_Coll_All


@mode int,
@StDate datetime,
@EdDate datetime
As
set nocount on

if @mode=1
begin

--drop table #Adv_Coll1

--drop table #tmp1
--drop table #tmp3
--drop table #Pre_tmp
--drop table #Pre_due_tmp
--drop table #tmp2
--drop table #tmp4
--drop table #tmp
--drop table #tmpp11
--drop table #tmp1
--drop table #Adv_Coll
--drop table #Adv_Coll2



CREATE TABLE [#Adv_Coll1] (
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[adv_coll1] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL,
	[u_name] [varchar] (1000) NULL ,
	[pre_coll] [money] NULL,
	[pre_due_coll] [money] NULL 

) ON [PRIMARY]



------due->>>>>>>>>>>>>>>>>>>>---------------------
select c.pat_id,due=isnull(sum(c.adv),0),
doc_name=isnull((select doc_name from doctor_info e where
e.refer_code=a.refer_code) ,(select doc_name from doctor_info_new f where
f.pat_id=a.pat_id))

into #tmp3 from pat_info_sub2 c,pat_info_main a
where c.type='due' 
and c.adv<>0
and a.pat_id=c.pat_id
and c.dt between @StDate and @EdDate
group by c.pat_id,
a.refer_code,a.pat_id

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,0,0,0,adv=due,doc_name,'0',0,0 from #tmp3
---->>>>>>>>>>>>>>>>>>>>>>>>
--select*from #Adv_Coll1
--select*from #tmp3
--select*from pat_info_sub2

---*>*>*>*>*PREVIOUS COLLECTION->>>>>>>>>>

select t.pat_id,pre_coll=sum(c.adv),t.doc_name into #Pre_tmp from pat_info_sub2 c,#tmp3 t where
c.pat_id=t.pat_id and dt<@StDate group by t.pat_id,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,0,0,0,0,doc_name,'0',pre_coll,0 from #Pre_tmp
--select*from #Adv_Coll1
--select*from #Pre_tmp
---*<*<*<*<*<--END PREVIOUS COLLECTION--<<<<--


---=>=>=>=>=PREVIOUS DUE COLLECTION->>>>>>>>>>

select t.pat_id,pre_due_coll=sum(c.adv),t.doc_name into #Pre_due_tmp from pat_info_sub2 c,#tmp3 t 
where c.pat_id=t.pat_id and c.type='due' and dt<@StDate group by t.pat_id,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,0,0,0,0,doc_name,'0',0,pre_due_coll from #Pre_due_tmp
--select*from #Adv_Coll1
--select*from #Pre_tmp
---=<=<=<=<=<--END PREVIOUS DUE COLLECTION--<<<<--


----ADVANCE collection>>>>>>>>>>>>>---------
select c.pat_id,adv=isnull(sum(c.adv),0),t.doc_name
into #tmp2 from pat_info_sub2 c,#tmp3 t where 
c.type='adv' and  t.pat_id=c.pat_id 
group by c.pat_id,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,0,adv=sum(adv),0,0,doc_name,'0',0,0 from #tmp2 group by pat_id,doc_name


--<<<<<<<<--Advance collection---------
---collect_fee->>>>>--------------------
select c.pat_id,collect_fee=isnull(sum(c.collect_fee),0),t.doc_name
into #tmp4
from pat_info_sub2 c,#tmp3 t
where t.pat_id=c.pat_id
group by c.pat_id,c.collect_fee,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,0,0,collect_fee,0,0,0,doc_name,'0',0,0 from #tmp4

--<<<<<--------------------------------

-----test rate->>>>>>>>--------------------------------------------
select b.pat_id,test_rate=sum(b.test_rate),t.doc_name
into #tmp 
from pat_info_sub1 b,#tmp3 t
where b.pat_id=t.pat_id
group by b.pat_id,t.doc_name

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,test_rate,0,0,0,0,0,0,doc_name,'0',0,0 from #tmp

-----test rate <<<<<<<<<<<<--------------- 
-->>>-discount----------------------------
---DISCOUNT>>>>>>>>>>>>>-------------
select d.pat_id,disc=sum(d.disc)
into #tmpp11
from pat_info_sub3 d,#tmp3 t
where d.pat_id=t.pat_id
group by d.pat_id
---<<------

---VAT>>>>>>>>>>>>>
select a.pat_id,t.doc_name,a.vat_amt,d.disc
into #tmp1
from pat_info_main a,
#tmpp11 d,pat_info_sub2 c,#tmp3 t
where a.pat_id=t.pat_id
and a.pat_id=d.pat_id and a.pat_id=c.pat_id
group by a.pat_id,a.refer_code,t.doc_name,
a.vat_amt,d.disc

insert into #Adv_Coll1(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,0,vat_amt,disc,0,0,0,0,doc_name,'0',0,0 from #tmp1

---<<<<<<<<<<<VAT,DISCOUNT--------


---fo create Auto_No------->>>>>>>>>>>>>>>
CREATE TABLE [#Adv_Coll] (
	[sl_no] [int] IDENTITY (1, 1) NOT NULL ,
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[adv_coll1] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL, 
	[u_name] [varchar] (1000) NULL,
	[pre_coll] [money] NULL,
	[pre_due_coll] [money] NULL

) ON [PRIMARY]

insert into #adv_coll(pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,adv_coll1,due_coll,doc_name,u_name,pre_coll,pre_due_coll)
select pat_id,test_rate=sum(test_rate),VAT_amt=sum(VAT_amt),
disc=sum(disc),coll_fee=sum(coll_fee),adv_coll=sum(adv_coll),adv_coll1=sum(adv_coll1),
due_coll=sum(due_coll),doc_name,'0',pre_coll=SUM(pre_coll),pre_due_coll=sum(pre_due_coll)
from  #Adv_Coll1
group by pat_id,doc_name
--<<<<<<<<<<---------------------

----final->>>>>>>>>>>>--------------

select t.pat_id,t.test_rate,t.VAT_amt,
total=(t.test_rate+t.VAT_amt+t.coll_fee),
t.pre_coll,Due_Amt=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc-t.pre_coll),
farther_due=(t.test_rate+t.VAT_amt+t.coll_fee-t.disc-t.adv_coll-t.due_coll-t.adv_coll1),
t.disc,t.coll_fee,t.adv_coll,t.due_coll,t.doc_name,a.pat_id1,t.pre_due_coll
into #Adv_Coll2 from #adv_coll t,pat_info_main a 
where t.pat_id=a.pat_id order by t.pat_id

update #Adv_Coll2 set pat_id1=pat_id where pat_id1=''

update #Adv_Coll2 set pre_coll=adv_coll where pre_coll=0 and adv_coll<>0

select pat_id,test_rate,VAT_amt,total,pre_coll,Due_Amt=(total-pre_coll),
farther_due,disc,coll_fee,adv_coll,due_coll,doc_name,
pat_id1,pre_due_coll from #Adv_Coll2


--select*from #Adv_Coll2





end


set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE       PROCEDURE Due_Doc_Pat


@refer_code varchar(10),
@StDate datetime,
@EdDate datetime

AS

set nocount on
---set @refer_code='001'
---set @StDate='2001-05-11 15:16:01.463'
---set @EdDate='2002-06-11 15:16:01.463'
select a.pat_id,a.pat_name,a.addr,a.phone,a.refer_code,
f.doc_name,a.vat_amt,disc=sum(b.disc) 
into #pat_info_main from pat_info_main a,pat_info_sub3 b,doctor_info f
where a.refer_code=@refer_code 
and a.refer_code=f.refer_code
and a.pat_id=b.pat_id 
and a.dt between @StDate and @EdDate
group by a.pat_id,a.pat_name,a.addr,a.phone,a.refer_code,
f.doc_name,a.vat_amt

----part 2----------------
--drop table #pat_info_sub
select b.pat_id,test_rate=sum(b.test_rate) 
into #pat_info_sub from pat_info_sub1 b,#pat_info_main a
where b.pat_id=a.pat_id
group by b.pat_id
--select * from #pat_info_sub

----part 3------------


--drop table #temp3
select c.pat_id,adv=sum(c.adv),collect_fee=sum(c.collect_fee)
into #temp3 from pat_info_sub2 c,#pat_info_main a
where a.pat_id=c.pat_id
group by c.pat_id

--select*from #temp3

------part 4---------------
--drop table #tmp4
select a.pat_id,a.pat_name,a.addr,a.phone,a.refer_code,
a.doc_name,a.vat_amt,a.disc,b.test_rate,t.adv,t.collect_fee
into #tmp4 from #pat_info_main a,#pat_info_sub b,#temp3 t
where a.pat_id=b.pat_id and t.pat_id=a.pat_id

--drop table #tmp5
select pat_id,pat_name,addr,phone,refer_code,doc_name,
vat_amt,disc,test_rate,adv,collect_fee,Tot_Bill=(test_rate+vat_amt+collect_fee-disc),
due=((test_rate+vat_amt+collect_fee-disc)-adv)
into #tmp5 from #tmp4

--select * from #tmp5 where due > 0
select t.pat_id,t.pat_name,t.addr,t.phone,t.refer_code,t.doc_name,t.vat_amt,
t.disc,t.test_rate,t.adv,t.collect_fee,t.Tot_Bill,t.due,a.pat_id1,Actual_Tot_Bill=(t.Tot_Bill-t.vat_amt)
into #tmp6 from #tmp5 t,pat_info_main a where a.pat_id=t.pat_id and t.due>0

update #tmp6 set pat_id1=pat_id where pat_id1=''
select*from #tmp6 order by pat_id

set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO










/****** Object:  Stored Procedure dbo.Due_Pat    Script Date: 03/04/2002 4:10:14 PM ******/
CREATE        PROCEDURE [Due_Pat] 

@StDate as datetime,
@EdDate as datetime

AS
set nocount on
--drop table #tmp
--drop table #tmp1
--drop table #tmp2
--drop table #tmp2N
--drop table #tmp3
--drop table #tmp3N
--drop table #tmp4
--drop table #tmp4N
--drop table #adv_coll
--drop table #Adv_Coll1
--drop table #tmp_tmp
--drop table #123
--drop table #due_coll


--declare @StDate as datetime
--declare @EdDate as datetime

--set @StDate='2002-03-01 11:30:00.000'
--set @EdDate='2002-05-09 11:34:00.000'
delete from Pat_due_Coll

CREATE TABLE [#Adv_Coll1] (
	[dt] [datetime] NULL ,
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL ,
	[pat_name] [varchar] (60) NULL ,
	[addr] [varchar] (100) NULL ,
	[phone] [varchar] (25) NULL 

) ON [PRIMARY]
--select a.pat_name,a.addr,a.phone from pat_info_main a,pat_info_sub1 b where a.pat_id=b.pat_id
-----***Test Rate----

select b.tmp_dt,b.pat_id,test_rate=sum(b.test_rate),
doc_name=isnull((select doc_name from doctor_info e where
e.refer_code=a.refer_code) ,(select doc_name from doctor_info_new f where
f.pat_id=a.pat_id)),
a.pat_name,a.addr,a.phone 
into #tmp from pat_info_sub1 b,pat_info_main a
where a.pat_id=b.pat_id 
and b.dt1 between @StDate and @EdDate 
group by b.pat_id,b.tmp_dt,
a.refer_code,a.pat_id,a.pat_name,a.addr,a.phone

--select*from pat_info_sub1
--select f.doc_name from doctor_info f,pat_info_main a where a.refer_code=f.refer_code
--select*from #tmp
insert into #Adv_Coll1(dt,pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,pat_name,addr,phone)
select tmp_dt,pat_id,test_rate,0,0,0,0,0,doc_name,pat_name,addr,phone from #tmp

-----***end----

----->>>for PAT_ID only-->>
--drop table #tmpp1
select distinct c.pat_id
into #tmpp1 from pat_info_sub2 c
where c.dt1 between @StDate and @EdDate
--<<END PAT_ID only

select a.tmp_dt,a.pat_id,a.refer_code,
doc_name=isnull((select doc_name from doctor_info e where
e.refer_code=a.refer_code) ,(select doc_name from doctor_info_new f where
f.pat_id=a.pat_id))
,a.vat_amt,disc=sum(d.disc),a.pat_name,a.addr,a.phone
into #tmp1
from pat_info_main a,
pat_info_sub3 d,#tmpp1 c
where a.pat_id=d.pat_id and a.pat_id=c.pat_id
--and c.dt1 between @StDate and @EdDate
group by a.tmp_dt,a.pat_id,a.refer_code,
a.vat_amt,
a.pat_name,a.addr,a.phone

insert into #Adv_Coll1(dt,pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,pat_name,addr,phone)
select tmp_dt,pat_id,0,vat_amt,disc,0,0,0,doc_name,pat_name,addr,phone from #tmp1

-----adv----------------------------------
select tmp_dt,pat_id,adv=isnull((select sum(adv) where type='adv'),0)
into #tmp2
from pat_info_sub2
where dt1 between @StDate and @EdDate
group by pat_id,type,tmp_dt

select t.tmp_dt,t.pat_id,t.adv,
doc_name=isnull((select doc_name from doctor_info e where
e.refer_code=a.refer_code) ,(select doc_name from doctor_info_new f where
f.pat_id=a.pat_id)),a.pat_name,a.addr,a.phone
into #tmp2N from #tmp2 t,
pat_info_main a where t.pat_id=a.pat_id 

insert into #Adv_Coll1(dt,pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,pat_name,addr,phone)
select tmp_dt,pat_id,0,0,0,0,adv=sum(adv),0,doc_name,pat_name,addr,phone from #tmp2N 
group by pat_id,tmp_dt,doc_name,pat_name,addr,phone

------------------------------------
----due-----------------------------------
select tmp_dt,pat_id,due=isnull((select sum(adv) where type='due'),0)
into #tmp3
from pat_info_sub2
where dt1 between @StDate and @EdDate
--where dt1<=@StDate and dt1=@EdDate
group by pat_id,type,tmp_dt



select t.tmp_dt,t.pat_id,t.due,
doc_name=isnull((select doc_name from doctor_info e where
e.refer_code=a.refer_code) ,(select doc_name from doctor_info_new f where
f.pat_id=a.pat_id)),a.pat_name,a.addr,a.phone
into #tmp3N from #tmp3 t,
pat_info_main a where t.pat_id=a.pat_id 

insert into #Adv_Coll1(dt,pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,pat_name,addr,phone)
select tmp_dt,pat_id,0,0,0,0,0,adv=sum(due),doc_name,pat_name,addr,phone 
from #tmp3N group by pat_id,tmp_dt,doc_name,pat_name,addr,phone
------------------------------------

----collect_fee-----------------------------------
select tmp_dt,pat_id,collect_fee
into #tmp4
from pat_info_sub2
where dt1 between @StDate and @EdDate
group by pat_id,collect_fee,tmp_dt

select t.tmp_dt,t.pat_id,t.collect_fee,f.doc_name,a.pat_name,a.addr,a.phone
into #tmp4N from #tmp4 t,doctor_info f,pat_info_main a where t.pat_id=a.pat_id 
and a.refer_code=f.refer_code


insert into #Adv_Coll1(dt,pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,pat_name,addr,phone)
select tmp_dt,pat_id,0,0,0,collect_fee=(sum(collect_fee)),0,0,doc_name,pat_name,addr,phone 
from #tmp4N group by pat_id,tmp_dt,doc_name,pat_name,addr,phone
------------------------------------

select pat_id,doc_name into #123
from #adv_coll1 group by pat_id,doc_name

select pat_id,doc_name into #tmp_tmp from #123 where doc_name<>'' group by pat_id,doc_name



CREATE TABLE [#Adv_Coll] (
	[sl_no] [int] IDENTITY (1, 1) NOT NULL ,
	[dt1] [datetime] NULL ,
	[pat_id] [int] NULL ,
	[test_rate] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[doc_name] [varchar] (60) NULL ,
	[pat_name] [varchar] (60) NULL ,
	[addr] [varchar] (100) NULL ,
	[phone] [varchar] (25) NULL 

) ON [PRIMARY]

insert into #adv_coll(dt1,pat_id,test_rate,VAT_amt,disc,coll_fee,adv_coll,due_coll,doc_name,pat_name,addr,phone)
select a.dt,a.pat_id,test_rate=sum(a.test_rate),VAT_amt=sum(a.VAT_amt),
disc=sum(a.disc),coll_fee=sum(a.coll_fee),
adv_coll=sum(a.adv_coll),due_coll=sum(a.due_coll),b.doc_name,a.pat_name,a.addr,a.phone
from #adv_coll1 a, #tmp_tmp b
where a.pat_id=b.pat_id
group by a.pat_id,b.doc_name,a.dt,a.pat_name,a.addr,a.phone

select sl_no,dt1,pat_id,test_rate,VAT_amt,
total=(test_rate+VAT_amt+coll_fee-disc),due_amt=((test_rate+VAT_amt+coll_fee)-(disc+adv_coll+due_coll)),
disc,coll_fee,adv_coll,due_coll,doc_name,pat_name,addr,phone
into #due_coll 
from #adv_coll


--select * from #due_coll
--where due_amt>0


select t.sl_no,t.dt1,t.pat_id,t.test_rate,t.VAT_amt,t.total,t.due_amt,
t.disc,t.coll_fee,t.adv_coll,t.due_coll,t.doc_name,t.pat_name,t.addr,t.phone,a.pat_id1
into #Adv_Coll2 from #due_coll t,pat_info_main a
where t.pat_id=a.pat_id and t.due_amt>0

update #Adv_Coll2 set pat_id1=pat_id where pat_id1=''

insert into Pat_due_Coll
select dt1,pat_id,test_rate,VAT_amt,total,due_amt,disc,
coll_fee,adv_coll,due_coll,doc_name,pat_name,addr,
phone,pat_id1 from #Adv_Coll2

select msg='Process Completed'

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE  PROCEDURE FLUSH_Com

@Status int,
@refer_code varchar(20)

AS


if @Status=1

begin

	select type Type,refer_code Doctor_ID,
	Doctor_Name=(select doc_name from doctor_info where doctor_info.refer_code=commission_per.refer_code),comm_per Commission 
	from commission_per where refer_code=@refer_code order by refer_code

end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE    PROCEDURE Flush_Test_Result

@Status int,
@type varchar(4),
@m_code varchar(2),
@s_code varchar(3)

AS

if @Status=1

begin

	select distinct *
	from test_result where type=@type
	and m_code=@m_code and s_code=@s_code

end

if @Status=2

begin

	select distinct *
	from test_result where type=@type
	and m_code=@m_code

end

--select *from test_result



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




CREATE  PROCEDURE Font_IU

@screen_name varchar(20),
@font_name varchar(30),
@font_type int
As
set nocount on
if exists(select*from font where screen_name=@screen_name)
	begin

	UPDATE Font SET font_name=@font_name,font_type=@font_type
	WHERE screen_name=@screen_name

	end
else
	begin

	INSERT INTO Font(screen_name,font_name,font_type)
	VALUES(@screen_name,@font_name,@font_type)

	end


select message='Successfully Operated'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE   PROC GetTestName
@m_code varchar(2),
@pat_id varchar(10)
AS

	select  * from test_result where  m_code=@m_code and
	s_code in(select s_code from pat_info_sub1 where pat_id=@pat_id and  m_code=@m_code)





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE   PROCEDURE Ins_Into_VAT 

@status int,
@pat_id  int

AS

if @status=1
begin
--part 1

	delete from pat_info_main_VAT where pat_id=@pat_id
	insert into pat_info_main_VAT(pat_id,pat_name,sex,age,refer_code,addr,phone,fax,email,uid,dt,vat_per,vat_amt,
	booth,tmp_dt,dt1,refer_type,pat_id1)
	select pat_id,pat_name,sex,age,refer_code,addr,phone,fax,email,uid,dt,vat_per,vat_amt,
	booth,tmp_dt,dt1,refer_type,pat_id1  from pat_info_main where pat_id=@pat_id

--end part 1
--part 2--

	delete from pat_info_sub1_VAT where pat_id=@pat_id
	insert into pat_info_sub1_VAT(pat_id,m_code,s_code,test_rate,delv_dt,type,
	uid,dt,tmp_dt,dt1)
	select pat_id,m_code,s_code,test_rate,delv_dt,type,
	uid,dt,tmp_dt,dt1 from pat_info_sub1 where pat_id=@pat_id

---end part 2--
---part-3

	delete from pat_info_sub2_VAT where pat_id=@pat_id
	insert into pat_info_sub2_VAT(pat_id,adv,uid,dt,collect_fee,type,tmp_dt,dt1)
	select pat_id,adv,uid,dt,collect_fee,type,tmp_dt,dt1 from pat_info_sub2 where pat_id=@pat_id

--end part 3--
--part 4
	delete from pat_info_sub3_VAT where pat_id=@pat_id
	insert into pat_info_sub3_VAT(pat_id,disc,paid,uid,dt,tmp_dt,dt1)
	select pat_id,disc,paid,uid,dt,tmp_dt,dt1 from pat_info_sub3 where pat_id=@pat_id

--part 4
end

if @status=2
begin

	delete from pat_info_main_VAT where pat_id=@pat_id
	delete from pat_info_sub1_VAT where pat_id=@pat_id
	delete from pat_info_sub2_VAT where pat_id=@pat_id
	delete from pat_info_sub3_VAT where pat_id=@pat_id

end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Inv_NO_SELECT    Script Date: 29/04/2002 12:40:12 AM ******/

/****** Object:  Stored Procedure dbo.Inv_NO_SELECT    Script Date: 24/04/2002 05:16:16 PM ******/

/****** Object:  Stored Procedure dbo.Inv_NO_SELECT    Script Date: 20/04/2002 08:35:24 PM ******/
CREATE PROCEDURE Inv_NO_SELECT
@status char(2),
@Inv_no varchar(50),
@Item_code varchar(20)

AS



if @status=1
begin


	select * from stock_in where inv_no=@Inv_no and Item_code=@Item_code

end

if @status=2
begin


	select * from stock_out where out_no=@Inv_no and Item_code=@Item_code

end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Item_Info_IUD    Script Date: 29/04/2002 12:40:12 AM ******/

/****** Object:  Stored Procedure dbo.Item_Info_IUD    Script Date: 24/04/2002 05:16:16 PM ******/

/****** Object:  Stored Procedure dbo.Item_Info_IUD    Script Date: 20/04/2002 08:35:24 PM ******/
CREATE PROCEDURE Item_Info_IUD


@Status varchar (1),
@item_code varchar (10),
@item_name varchar (50)

 AS


if @Status='I'
begin
	insert into item_info(item_code,item_name)
	values(@item_code,@item_name)
end

if @Status='U'
begin
	update item_info set item_code=@item_code,item_name=@item_name where item_code=@item_code
end

if @Status='D'
begin
	delete from item_info where item_code=@item_code
end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




/****** Object:  Stored Procedure dbo.Leave_As_Cash_IUD    Script Date: 29/04/2002 12:40:12 AM ******/

/****** Object:  Stored Procedure dbo.Leave_As_Cash_IUD    Script Date: 24/04/2002 05:16:15 PM ******/
CREATE PROCEDURE Leave_As_Cash_IUD

@status char(2),
@emp_id varchar (10),
@leave_type varchar (15),
@leave money,
@cash_date datetime 

AS

if @status='I'
begin
	insert into Leave_as_Cash(emp_id,leave_type,leave,cash_date)
	values(@emp_id,@leave_type,@leave,@cash_date)
end

if @status='U'
begin
	update Leave_as_Cash set emp_id=@emp_id,leave_type=@leave_type,leave=@leave,cash_date=@cash_date 
	where emp_id=@emp_id and leave_type=@leave_type and cash_date=@cash_date

end

if @status='D'
begin
	delete from Leave_as_Cash where emp_id=@emp_id and leave_type=@leave_type and cash_date=@cash_date
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE            PROCEDURE Leave_Balance

@status int,
@emp_id varchar(10)

AS

set nocount on

if @status=1
begin

/*
declare @emp_name varchar(50)
set @emp_name=(select emp_name from emp_info where emp_id=@emp_id)


CREATE TABLE [#getdate1] (
	[Curr_Date] [datetime] NULL ,
)
insert into #getdate1(Curr_Date)
select getdate()

--drop table #getdate2
select Curr_Date into #getdate2 from #getdate1

--drop table #getdate3
select cast(Curr_Date as varchar (11)) StrDate into #getdate3 from #getdate2

--drop table #getdate4
select Cryear=SUBSTRING(StrDate,8,11) into #getdate4 from #getdate3
-------------------------------------------------------------------------
--drop table #leave1
SELECT Emp_ID=@emp_id,Leave_st_Date=cast(Leave_st_Date as varchar(11)),
Leave_ed_Date=cast(Leave_ed_Date as varchar(11)),
Leave_Type,Total_Leave into #leave1 FROM leave where emp_id=@Emp_ID

--select * from #leave1 order by leave_st_date


--drop table #leave2
select Emp_ID=@emp_id,StYear=SUBSTRING(Leave_st_Date,8,11),EdYear=SUBSTRING(Leave_ed_Date,8,11),
Leave_Type,Total_Leave into #leave2 from #leave1
--select*from #leave1
--drop table #leave3
select Emp_ID=@emp_id,emp_name=@emp_name,
a.Leave_Type,Tot_Availed_Leave=sum(a.Total_Leave),b.Celing,
Balance=(b.Celing-sum(a.Total_Leave))
into #leave3 from #leave2 a,lv_setup b,#getdate4 c,emp_info e
where a.emp_id=e.emp_id and a.Leave_Type=b.Leave_Type and a.StYear=c.Cryear and a.EdYear=c.Cryear
group by e.Emp_Name,a.StYear,a.EdYear,
a.Leave_Type,b.Leave_Type,b.Celing
---------------------------------------------
--select*from #leave3


		select* into #leave4 from #leave3
		union
--		select Emp_id=(select Emp_id from #leave3 group by Emp_id),
		select Emp_id=@Emp_id,
		Emp_name=@emp_name,
		a.Leave_type,
		Tot_Availed_Leave=0,
		a.Celing,
		Balance=a.Celing
		from lv_setup a
		where not exists (select * from #leave3 b where a.Leave_type=b.Leave_type)


declare @leave_st_date datetime
declare @leave_ed_date datetime
set @leave_st_date=(select leave_st_date from leave where track_id=(select track_id=max(track_id) from leave where emp_id=@emp_id))
set @leave_ed_date=(select leave_ed_date from leave where track_id=(select track_id=max(track_id) from leave where emp_id=@emp_id))

select Emp_ID,emp_name,Leave_Type,Tot_Availed_Leave,
Celing,Balance,leave_st_date=@leave_st_date,
leave_ed_date=@leave_ed_date from #leave4
*/


declare @emp_name varchar(50)
set @emp_name=(select emp_name from emp_info where emp_id=@emp_id)


CREATE TABLE [#getdate1] (
	[Curr_Date] [datetime] NULL ,
)
insert into #getdate1(Curr_Date)
select getdate()

--drop table #getdate2
select Curr_Date into #getdate2 from #getdate1

--drop table #getdate3
select cast(Curr_Date as varchar (11)) StrDate into #getdate3 from #getdate2

--drop table #getdate4
select Cryear=SUBSTRING(StrDate,8,11) into #getdate4 from #getdate3
-------------------------------------------------------------------------
--drop table #leave1
SELECT Emp_ID=@emp_id,Leave_st_Date=cast(Leave_st_Date as varchar(11)),
Leave_ed_Date=cast(Leave_ed_Date as varchar(11)),
Leave_Type,Total_Leave into #leave1 FROM leave where emp_id=@Emp_ID

--select * from #leave1 order by leave_st_date


--drop table #leave2
select Emp_ID=@emp_id,StYear=SUBSTRING(Leave_st_Date,8,11),EdYear=SUBSTRING(Leave_ed_Date,8,11),
Leave_Type,Total_Leave into #leave2 from #leave1
--select*from #leave1
--drop table #leave3
select Emp_ID=@emp_id,emp_name=@emp_name,
a.Leave_Type,Tot_Availed_Leave=sum(a.Total_Leave),b.Celing,
Balance=(b.Celing-sum(a.Total_Leave))
into #leave3 from #leave2 a,lv_setup b,#getdate4 c,emp_info e
where a.emp_id=e.emp_id and a.Leave_Type=b.Leave_Type and a.StYear=c.Cryear and a.EdYear=c.Cryear
group by e.Emp_Name,a.StYear,a.EdYear,
a.Leave_Type,b.Leave_Type,b.Celing
---------------------------------------------
--select*from #leave3


		select* into #leave4 from #leave3
		union
--		select Emp_id=(select Emp_id from #leave3 group by Emp_id),
		select Emp_id=@Emp_id,
		Emp_name=@emp_name,
		a.Leave_type,
		Tot_Availed_Leave=0,
		a.Celing,
		Balance=a.Celing
		from lv_setup a
		where not exists (select * from #leave3 b where a.Leave_type=b.Leave_type)


declare @leave_st_date datetime
declare @leave_ed_date datetime
--set @leave_st_date=(select leave_st_date from leave where track_id=(select track_id=max(track_id) from leave where emp_id=@emp_id))
--set @leave_ed_date=(select leave_ed_date from leave where track_id=(select track_id=max(track_id) from leave where emp_id=@emp_id))

set @leave_st_date=(select max(leave_st_date) from leave where emp_id=@emp_id)
set @leave_ed_date=(select max(leave_ed_date) from leave where emp_id=@emp_id)

select Emp_ID,emp_name,Leave_Type,Tot_Availed_Leave,
Celing,Balance,leave_st_date=@leave_st_date,
leave_ed_date=@leave_ed_date into #leave5 from #leave4

--select*from #leave5
-->>UPDATE EARNED LEAVE CELING,BALANCE--->>>>>>-----------------------

declare @Cr_Year int
set @Cr_Year=(select cryear from #getdate4)

declare @styear1 varchar(11)
set @styear1=(select distinct cast(Lv_st_year as varchar(11)) from lv_setup)

declare @year_diff int
declare @styear varchar(4)
set @styear=((select distinct SUBSTRING(@styear1,8,11) from lv_setup))
set @year_diff=(@Cr_Year-@styear)+1
--select @year_diff 


CREATE TABLE #update
	(emp_id varchar(20),
	leave_type varchar(20),
	total_leave money,
	celing int,
	balance int
	)
insert into #update(emp_id,leave_type,total_leave,celing,balance)
		values(@emp_id,'earned',0,0,0)

declare @celing int
set @celing=(select (celing*@year_diff) from lv_setup where leave_type='earned')

insert into #update
select a.emp_id,a.leave_type,total_leave=sum(a.total_leave),
celing=(b.celing*@year_diff),balance=((b.celing*@year_diff)-sum(a.total_leave))
--into #update 
from leave a,lv_setup b where a.leave_type='earned' 
and a.leave_type=b.leave_type
and a.emp_id=@emp_id
group by a.emp_id,a.leave_type,b.celing

select emp_id,leave_type,total_leave=sum(total_leave),
celing=sum(celing),balance=sum(balance) 
into #update1 from #update group by emp_id,leave_type

declare @Earned_celing int
declare @Earned_Bal int

set @Earned_celing=(select celing from #update1)
set @Earned_Bal=(select balance from #update1)

update #leave5 set balance=@Earned_Bal where leave_type='earned'
update #leave5 set celing=@celing where leave_type='earned'
update #leave5 set tot_availed_leave=(select isnull(sum(total_leave),0) from leave where emp_id=@emp_id and leave_type='earned') where leave_type='earned'

select Emp_ID,emp_name,Leave_Type,Tot_Availed_Leave,Celing,Balance=(Celing-Tot_Availed_Leave),leave_st_date,leave_ed_date from #leave5

end

if @status=2
begin

	select*from item_info

end

if @status=3
begin

	SELECT s.out_no,s.emp_id,e.Emp_Name,s.item_code,
	i.item_name,s.item_qty Qty,s.issu_date,s.notes,s.u_id
	FROM stock_out s,emp_info e,item_info i
	where s.emp_id=@emp_id and s.emp_id=e.emp_id 
	and i.item_code=s.item_code
end

if @status=4
begin

	select Emp_ID,Emp_Name,join_date,Emp_Desig,Title,Salary,
	Sex,Age,Emp_Per_Add Present_Address,Emp_Pre_Add Permanent_Address,Emp_Phone,Emp_Email from emp_info

end

set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Leave_Balance1    Script Date: 29/04/2002 12:40:12 AM ******/

/****** Object:  Stored Procedure dbo.Leave_Balance1    Script Date: 24/04/2002 05:16:17 PM ******/

/****** Object:  Stored Procedure dbo.Leave_Balance1    Script Date: 20/04/2002 08:35:24 PM ******/
CREATE PROCEDURE Leave_Balance1

@status int
--@emp_id varchar(10)

AS

set nocount on

if @status=1
begin

--drop table #getdate1
CREATE TABLE [#getdate1] (
	[Curr_Date] [datetime] NULL ,
)
insert into #getdate1(Curr_Date)
select getdate()

--drop table #getdate2
select Curr_Date into #getdate2 from #getdate1

--drop table #getdate3
select cast(Curr_Date as varchar (11)) StrDate into #getdate3 from #getdate2

--drop table #getdate4
select Cryear=SUBSTRING(StrDate,8,11) into #getdate4 from #getdate3
-------------------------------------------------------------------------
--declare @Emp_ID as varchar(10)
--set @Emp_ID='001'
--drop table #leave1
SELECT Emp_ID,Leave_st_Date=cast(Leave_st_Date as varchar(11)),
Leave_ed_Date=cast(Leave_ed_Date as varchar(11)),
Leave_Type,Total_Leave into #leave1 FROM leave 
--where emp_id=@Emp_ID

--drop table #leave2
select Emp_ID,StYear=SUBSTRING(Leave_st_Date,8,11),EdYear=SUBSTRING(Leave_ed_Date,8,11),
Leave_Type,Total_Leave into #leave2 from #leave1

select a.Emp_ID,e.emp_name,
a.Leave_Type,Done_Leave=sum(a.Total_Leave),b.Celing,
Balance=(b.Celing-sum(a.Total_Leave))
from #leave2 a,lv_setup b,#getdate4 c,emp_info e
where a.emp_id=e.emp_id and a.Leave_Type=b.Leave_Type and a.StYear=c.Cryear and a.EdYear=c.Cryear
group by a.Emp_ID,e.Emp_Name,a.StYear,a.EdYear,
a.Leave_Type,b.Leave_Type,b.Celing
	

end

set nocount off








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE    PROC Make_Pat_ID1

@refer_type varchar(3)

As
set nocount on
--set @refer_type='1'
/*

---<<<Select Month and Year from GETDATE()
	declare @Tot_Date varchar(8)
	set @Tot_Date=(SELECT CONVERT(char(12), GETDATE(), 3))
	--------------
	declare @CrDay varchar(2)
	set @CrDay=(SELECT SUBSTRING(@Tot_Date,1,2) )
	--------------
	declare @CrMonth varchar(2)
	set @CrMonth=(SELECT SUBSTRING(@Tot_Date, 4,2) )
	--------------
	declare @CrYear varchar(2)
	set @CrYear=(SELECT SUBSTRING(@Tot_Date,7,2) )
	--select @CrMonth CrMonth,@CrYear CrYear
---<<<End Month and Year from GETDATE()

---SEARCH LAST PATIENT ID-->>

declare @pat_id1_Max varchar(50)
--drop table #tmp1
select pat_id,refer_type,pat_id1=isnull(pat_id1,'0'),pat_my
into #tmp1 from pat_info_main where pat_my=(@CrMonth+@CrYear) and refer_type=@refer_type

--select*from #tmp1

set @pat_id1_Max=(select isnull(max(substring(pat_id1,6,20)),'0') from #tmp1)
--select @pat_id1_Max pat_id1_Max

declare @pat_id1 varchar(50)

if @pat_id1_Max='0'
begin

	set @pat_id1=@CrMonth+@CrYear+'-'+'1'
	select pat_id1=@pat_id1,pat_my=(@CrMonth+@CrYear)


end
else
begin
	set @pat_id1_Max=@pat_id1_Max+1
	set @pat_id1=@CrMonth+@CrYear+'-'+@pat_id1_Max
	select pat_id1=@pat_id1,pat_my=(@CrMonth+@CrYear)
end
*/

---<<<Select Month and Year from GETDATE()
	declare @Tot_Date varchar(8)
	set @Tot_Date=(SELECT CONVERT(char(12), GETDATE(), 3))
	--------------
	declare @CrDay varchar(2)
	set @CrDay=(SELECT SUBSTRING(@Tot_Date,1,2) )
	--------------
	declare @CrMonth varchar(2)
	set @CrMonth=(SELECT SUBSTRING(@Tot_Date, 4,2) )
	--------------
	declare @CrYear varchar(2)
	set @CrYear=(SELECT SUBSTRING(@Tot_Date,7,2) )
	--select @CrMonth CrMonth,@CrYear CrYear
---<<<End Month and Year from GETDATE()

---SEARCH LAST PATIENT ID-->>

declare @pat_id1_Max varchar(50)

--drop table #tmp1
select pat_id,refer_type,pat_id1=isnull(pat_id1,'0'),pat_my
into #tmp1 from pat_info_main where pat_my=(@CrMonth+@CrYear) and refer_type=@refer_type

--select*from #tmp1


create table #tmp2
	( Dumy_ID int
	)

insert into #tmp2
select pp=(substring(pat_id1,6,20)) from #tmp1

set @pat_id1_Max=(select isnull(max(Dumy_ID),'0') from #tmp2)


declare @pat_id1 varchar(50)

if @pat_id1_Max='0'
begin

	set @pat_id1=@CrMonth+@CrYear+'-'+'1'
	select pat_id1=@pat_id1,pat_my=(@CrMonth+@CrYear)


end
else
begin
	set @pat_id1_Max=@pat_id1_Max+1
	
	set @pat_id1=@CrMonth+@CrYear+'-'+@pat_id1_Max
--	select @pat_id1_Max
	select pat_id1=@pat_id1,pat_my=(@CrMonth+@CrYear)
end



set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE     PROC Make_Pat_ID_U


@refer_type varchar(3)
As
set nocount on

declare @pat_id int
declare @pat_id1 varchar(50)
declare @CrID varchar(50)
--declare @refer_type varchar(3)
--set @refer_type='0'

---<<<Select Month and Year from GETDATE()
	declare @Tot_Date varchar(8)
	set @Tot_Date=(SELECT CONVERT(char(12), GETDATE(), 3))
	--------------
	declare @CrMonth varchar(2)
	set @CrMonth=(SELECT SUBSTRING(@Tot_Date, 4,2) )
	--------------
	declare @CrYear varchar(2)
	set @CrYear=(SELECT SUBSTRING(@Tot_Date,7,2) )
--	select @CrMonth CrMonth,@CrYear CrYear
---<<<End Month and Year from GETDATE()


set @pat_id=(select pat_id=max(pat_id) from pat_info_main where refer_type=@refer_type)

-->Select Month and Year from Last Patient------->>>>>>>>>>>>>>>>>>>>
declare @Pre_Tot_Date varchar(8)
set @Pre_Tot_Date=(select convert(char(12),(select dt1 from pat_info_main where pat_id=@pat_id),3))
--select @Pre_Tot_Date Pre_Tot_Date

declare @PreMonth varchar(2)
set @PreMonth=(SELECT SUBSTRING(@Pre_Tot_Date, 4,2) )
--------------
declare @PreYear varchar(2)
set @PreYear=(SELECT SUBSTRING(@Pre_Tot_Date,7,2) )
--select @PreMonth PreMonth,@PreYear PreYear
--<<<<<<<--Select Month and Year from Last Patient---<<<<<<<<<

if @PreMonth=@CrMonth and @PreYear=@CrYear
	begin

	set @CrID=(select pat_id1=(isnull(pat_id1,'0')) from pat_info_main where pat_id=@pat_id)

	set @pat_id1=(SELECT SUBSTRING(@CrID,6,20) )
	set @pat_id1=@pat_id1+1
	select pat_id1=(@CrMonth+@CrYear+'-'+@pat_id1),pat_my=(@CrMonth+@CrYear)

	end

if @PreMonth<>@CrMonth or @PreYear<>@CrYear
	begin

	set @CrID='0'
	set @pat_id1=(SELECT SUBSTRING(@CrID,6,20) )
	set @pat_id1=@pat_id1+1
	select pat_id1=(@CrMonth+@CrYear+'-'+@pat_id1),pat_my=(@CrMonth+@CrYear)

	end

set nocount off



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




CREATE  PROCEDURE New_Doc_Select

@Status int,
@refer_code varchar(100),
@uid varchar(15)

AS

if @Status=1
begin
	select * from Doctor_Info_new where pat_id=@refer_code and uid=@uid
end

if @Status=2
begin
	select pat_id,doc_name,addr,phone,fax,email,uid,doc_date 
	from doctor_info_new where pat_id='0' and uid=@uid
end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




CREATE  PROCEDURE PAT_INFO_MAIN_U
@status char(2),
@pat_id int
AS

if @status='U'
begin
	update pat_info_main set refer_code='' where pat_id=@pat_id

end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE             PROCEDURE Pat_Info_SELECT

@Status int,
@pat_id int

 AS

if @Status=1

begin
	select pat_id,pat_name,sex,age,refer_code,addr,phone,fax,email,uid,dt,
	vat_per,vat_amt,booth,tmp_dt,dt1,refer_type,pat_id1=isnull(pat_id1,''),pat_my from pat_info_main where pat_id=@pat_id

end

if @Status=2

begin
	select*from pat_info_sub2 where pat_id=@pat_id
end

if @Status=3

begin
	select*from pat_info_sub3 where pat_id=@pat_id

end

if @Status=4

begin
          delete from Pat_info_main where pat_id=@pat_id
end

if @Status=5

begin
	select*from pat_info_sub1 where pat_id=@pat_id

end

if @Status=6

begin
	select doc_name,addr from doctor_info_new where pat_id=@pat_id

end

if @Status=7

begin
	select refer_code from pat_info_main where pat_id=@pat_id

end

if @Status=8

begin
          delete from doctor_info_new where pat_id=@pat_id
end
if @Status=9
begin
	select pat_id,doc_name,addr,phone,fax,email,uid,doc_date from doctor_info_new where pat_id='0'
end

if @Status=10
begin


	declare @pat_id1 varchar(50)
	set @pat_id1=(select pat_id1 from pat_info_main where pat_id=@pat_id)

	if @pat_id1=''
	begin
		set @pat_id1=@pat_id
	end

	select b.pat_id,pat_id1=@pat_id1,a.pat_name,b.adv,b.collect_fee,b.unique_id
	from pat_info_sub2 b,pat_info_main a 
	where b.pat_id=@pat_id  and a.pat_id=b.pat_id order by b.unique_id


end
if @Status=11
begin
	select disc=sum(disc),paid=sum(paid) from pat_info_sub3 where pat_id=@pat_id
end
if @Status=12
begin

--	select b.pat_id,a.pat_name,b.disc,b.track_id
--	from pat_info_sub3 b,pat_info_main a 
--	where b.pat_id=@pat_id  and a.pat_id=b.pat_id
--	order by b.track_id


	declare @pat_id2 varchar(50)
	set @pat_id2=(select pat_id1 from pat_info_main where pat_id=@pat_id)

	if @pat_id2=''
	begin
		set @pat_id2=@pat_id
	end


	select b.pat_id,pat_id1=@pat_id2,a.pat_name,b.disc,b.track_id
	from pat_info_sub3 b,pat_info_main a 
	where b.pat_id=@pat_id  and a.pat_id=b.pat_id
	order by b.track_id




end



if @Status=13
begin


	declare @pat_id3 varchar(50)
	set @pat_id3=(select pat_id1 from pat_info_main where pat_id=@pat_id)

	if @pat_id3=''
	begin
		set @pat_id3=@pat_id
	end

	select distinct b.pat_id,pat_id1=@pat_id3,a.pat_name,
	b.m_code,b.s_code,b.type,b.unique_id
	from pat_info_sub1 b,pat_info_main a
	where b.pat_id=@pat_id  and a.pat_id=b.pat_id order by b.unique_id

end

if @Status=14
begin



	select distinct b.pat_id,pat_id1,a.pat_name,
	b.m_code,b.s_code,b.type,b.unique_id
	from pat_info_sub1 b,pat_info_main a
	where a.pat_id=b.pat_id and b.type='' and b.dt1>'2003-06-01 07:53:00.000'
	order by b.unique_id

end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROCEDURE Pat_Info_SELECT1

@Status int,
@pat_id varchar(50)

 AS

if @Status=1

begin
	select pat_id,pat_name,sex,age,refer_code,addr,phone,fax,email,uid,dt,
	vat_per,vat_amt,booth,tmp_dt,dt1,refer_type,pat_id1=isnull(pat_id1,'') from pat_info_main where pat_id=@pat_id

end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Pat_Info_SELECT_VAT    Script Date: 14/11/2001 12:26:01 AM ******/
CREATE PROCEDURE Pat_Info_SELECT_VAT

@Status int,
@pat_id varchar(10)

 AS
   
if @Status='1'

begin
	select*from pat_info_main_VAT where pat_id=@pat_id

end

if @Status='2'

begin
	select*from pat_info_sub2_VAT where pat_id=@pat_id
end

if @Status='3'

begin
	select*from pat_info_sub3_VAT where pat_id=@pat_id
end

if @Status=4

begin
          delete from Pat_info_main_VAT where pat_id=@pat_id
end

if @Status=5

begin
	select*from pat_info_sub1_VAT where pat_id=@pat_id

end

if @Status=6

begin
	select doc_name,addr from doctor_info_new where pat_id=@pat_id

end

if @Status=7

begin
	select refer_code from pat_info_main_VAT where pat_id=@pat_id

end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Pat_Info_Sub1_Delete    Script Date: 08/11/2001 6:17:02 PM ******/
/****** Object:  Stored Procedure dbo.Pat_Info_Sub1_Delete    Script Date: 26/09/2001 9:06:45 AM ******/
CREATE PROCEDURE Pat_Info_Sub1_Delete

@Status int,
@pat_id int

AS

if @Status=1

begin

	delete from pat_info_sub1 where pat_id=@pat_id

end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Pat_Info_Sub1_Delete1    Script Date: 26/09/2001 9:06:46 AM ******/
CREATE PROCEDURE Pat_Info_Sub1_Delete1

@Status int,
@Unique_ID int

 AS

if @Status=1

begin

	delete from pat_info_sub1 where unique_id=@Unique_ID
end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE Pat_Type


@refer_type as varchar (3),
@Stdate as datetime,
@Eddate as datetime

AS
set nocount on
--select a.pat_id,a.pat_name,
--doc_name=isnull((select doc_name from doctor_info c where
--c.refer_code=a.refer_code) ,(select doc_name from doctor_info_new d where
--d.pat_id=a.pat_id) ) from 
--pat_info_main a where a.refer_type=@refer_type
--and a.dt1 between @Stdate and @Eddate

select a.pat_id,a.pat_name,
doc_name=isnull((select doc_name from doctor_info c where
c.refer_code=a.refer_code) ,(select doc_name from doctor_info_new d where
d.pat_id=a.pat_id) ),a.pat_id1
into #tmp1 from pat_info_main a where a.refer_type=@refer_type
and a.dt1 between @Stdate and @Eddate

update #tmp1 set pat_id1=pat_id where pat_id1=''

select*from #tmp1

set nocount off



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.Pro_Auto    Script Date: 26/09/2001 9:06:46 AM ******/
CREATE PROCEDURE [Pro_Auto]
@status as tinyint
 AS
--declare @@type varchar(5)
declare @@autoNo varchar(8)
declare @@intno  numeric
declare @@str_no varchar(10)

--for PATIENT ID
if @status=1
begin
       select @@autoNo=autoNo  from auto_no where sl_no='1'		
        if @@autoNo=0 
           begin
	      set @@intno=1000001
           end
        else
            begin
                   set  @@intno=@@autoNo+1
            end     	
end

select xx=@@intno






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE Pro_Current_Stock
AS
set nocount on

select Item_code,no_of_box=sum(no_of_box),test_per_box,
tot_qty=(sum(no_of_box)*test_per_box)
into #stock from stock_in
group by Item_code,test_per_box
union
select Item_code,no_of_box=0,test_per_box=0,
tot_qty=-sum(item_qty)
from stock_out
group by Item_code

select a.Item_code,
Item_name=(select Item_name from item_info b where b.Item_code = a.item_code),
No_of_box=sum(a.no_of_box),Test_per_box=sum(a.test_per_box),
Balance_qty=sum(a.tot_qty) into #stock_Final from #stock a
group by a.Item_code

select Item_code,Item_name,No_of_box,Test_per_box,
Tot_Pur=No_of_box*Test_per_box,Tot_used=(No_of_box*Test_per_box)-Balance_qty,Balance_qty
from #stock_final

set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Pro_FLUSH    Script Date: 29/04/2002 12:40:12 AM ******/


/****** Object:  Stored Procedure dbo.Pro_FLUSH    Script Date: 20/10/2001 11:38:16 PM ******/

/****** Object:  Stored Procedure dbo.Pro_FLUSH    Script Date: 18/10/2001 6:18:24 PM ******/

CREATE  PROCEDURE Pro_FLUSH

@Status int,
@pat_id int

AS

if @Status=1
--for flush in gride from Doctor_Info Table
begin
	select refer_code,doc_name,addr,phone,fax,email,birth_date,marriage_date,dt from doctor_info order by refer_code
end

--flush from pat_info_sub2
if @Status=2
begin
	select pat_id,adv,dt,unique_id from pat_info_sub2 where pat_id=@pat_id
end

--flush from pat_info_sub2 for Advance SUM
if @Status=3
begin
	select adv_sum=isnull(sum(adv) ,0),Coll_sum=isnull(sum(collect_fee) ,0)  from Pat_Info_sub2 where pat_id=@pat_id
end

--for Report Entry Screen
if @Status=4
begin
           select a.m_code,a.s_code,
          (select s_name from test_info_sub b where a.s_code=b.s_code
          and a.m_code=b.m_code) as s_name 
         from pat_info_sub1 a where a.pat_id=@pat_id

end

if @Status=5
begin
	select * from Report_All where pat_id=@pat_id
end

-- flush from Report_All Table 
if @Status=6

begin
	select*from Report_All where pat_id=@pat_id
end

if @Status=7

begin
	--select m_code,refer_code,comm_per from commission_per
	select type Type,refer_code Doctor_ID,
	Doctor_Name=(select doc_name from doctor_info where doctor_info.refer_code=commission_per.refer_code),comm_per Commission 
	from commission_per order by refer_code

end

if @Status=8
--for flush in gride from Doctor_Info_new Table
begin
	select pat_id,doc_name,addr,phone,fax,email,doc_date,dt from doctor_info_new
end

if @Status=9
--for flush in gride from Doctor_Info_new Table
begin
	select doc_name,addr from Doctor_Info where refer_code=@pat_id
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



/****** Object:  Stored Procedure dbo.Pro_FLUSH1    Script Date: 20/01/2002 10:26:58 AM ******/
CREATE PROCEDURE Pro_FLUSH1

@Status int,
@refer_code varchar(10)

AS


if @Status=1
--for flush in gride from Doctor_Info_new Table
begin
	select doc_name,addr from Doctor_Info where refer_code=@refer_code
end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Pro_FLUSH_TN    Script Date: 29/04/2002 12:40:12 AM ******/

/****** Object:  Stored Procedure dbo.Pro_FLUSH_TN    Script Date: 24/04/2002 05:16:17 PM ******/


/****** Object:  Stored Procedure dbo.Pro_FLUSH_TN    Script Date: 08/11/2001 6:17:05 PM ******/
CREATE PROCEDURE Pro_FLUSH_TN

@Status int,
@m_code varchar(2),
@pat_id varchar(10)

AS

if @Status=1
--for flush in gride from 
begin
	select a.m_code,a.s_code,
(select s_name from test_info_sub b where a.s_code=b.s_code
          and a.m_code=b.m_code) as s_name 
         from pat_info_sub1 a where a.m_code=@m_code and a.pat_id=@pat_id
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Pro_FLUSH_VAT    Script Date: 14/11/2001 12:26:03 AM ******/
CREATE PROCEDURE Pro_FLUSH_VAT

@Status int,
@pat_id varchar(10)

AS

if @Status=1
--for flush in gride from Doctor_Info Table
begin
	select refer_code,doc_name,addr,phone,fax,email,dt from doctor_info
end

--flush from pat_info_sub2
if @Status=2
begin
	select pat_id,adv,dt,unique_id from pat_info_sub2_VAT where pat_id=@pat_id
end

--flush from pat_info_sub2 for Advance SUM
if @Status=3
begin
	select adv_sum=isnull(sum(adv) ,0)  from Pat_Info_sub2_VAT where pat_id=@pat_id
end

--for Report Entry Screen
if @Status=4
begin
           select a.m_code,a.s_code,
          (select s_name from test_info_sub b where a.s_code=b.s_code
          and a.m_code=b.m_code) as s_name 
         from pat_info_sub1_VAT a where a.pat_id=@pat_id

end

if @Status=5
begin
	select * from Report_All where pat_id=@pat_id
end

-- flush from Report_All Table 
if @Status=6

begin
	select*from Report_All where pat_id=@pat_id
end

if @Status=7

begin
	--select m_code,refer_code,comm_per from commission_per
	select type,
             refer_code,
             Doc_Name=(select doc_name from doctor_info where doctor_info.refer_code=commission_per.refer_code),comm_per from commission_per

end

if @Status=8
--for flush in gride from Doctor_Info_new Table
begin
	select pat_id,doc_name,addr,phone,fax,email,dt from doctor_info_new
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE Pro_Pur_Iss_Contrast
AS

Select Name_of_reagent=(select Item_name from Item_info where Item_code=a.Item_code),
a.pur_date,a.Item_code,a.No_of_box,a.Test_per_box,
Qty=(a.No_of_box*a.Test_per_box),b.Issu_date,b.Item_qty from stock_in a, stock_out b
where a.Item_code=b.Item_code



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE Pro_Stock_det
AS
set nocount on

CREATE TABLE [#Stk_Det] (
	[Sl_no] [int] IDENTITY (1, 1) NOT NULL  ,	
	[pur_date] [datetime] NULL ,
	[Name_of_sup] [varchar] (45) NULL ,
	[Name_of_reagent] [varchar] (45) NULL ,
	[Item_code] [varchar] (45) NULL ,
	[No_of_box] [int] NULL ,
	[Test_per_box] [int] NULL ,
	[Qty] [int] NULL ,
	[Exp_dt] [datetime] NULL ,
	[Pre_bal] [int] NULL 
)

insert into #stk_Det
Select a.pur_date,Name_of_sup=(select Sup_name from Sup_info where Sup_id=a.Sup_id),
Name_of_reagent=(select Item_name from Item_info where Item_code=a.Item_code),
a.Item_code,a.No_of_box,a.Test_per_box,
Qty=(a.No_of_box*a.Test_per_box),a.Exp_dt,Pre_bal=0 from stock_in a

declare @Counter as int
declare @No_of_Repeat as int
set @Counter=2
set @No_of_Repeat=(select count(*) from #stk_Det)

if @No_of_Repeat > 0
begin
	declare @Fdt as datetime
	declare @Tdt as datetime
	declare @Prebal as int
	declare @Out as int

	set @counter=2
St:
	set @Fdt=(Select pur_date from #Stk_det where Sl_no=(@counter-1))
	set @Tdt=(Select pur_date from #Stk_det where Sl_no=@counter)
	
	set @Out=isnull((select Sum(item_qty) from Stock_out where 
	Item_code = (Select item_code from #Stk_Det where sl_no=@counter)
	and issu_date between @Fdt and (@Tdt-1)),0)

	set @Prebal=(select (qty + Pre_bal) from #Stk_Det where sl_no=(@counter-1)) - @Out
	
	Update #Stk_Det set Pre_bal=@Prebal where Sl_no=@counter

	Set @counter=@counter+1
	if @counter <= @No_of_Repeat goto St

end

select * from #stk_Det

set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE PROCEDURE Pro_TOTrate_commission

@Status int

AS

if @Status=1

begin


select a.refer_code,d.doc_name,a.pat_id,b.pat_name,sum(c.test_rate) as tot_test_rate,a.commission,
paid=(select case count(paid) when 0 then 0 else sum(paid) end from commission_sub e where e.pat_id=a.pat_id)
from commission_main a,
pat_info_main b,
pat_info_sub1 c,
doctor_info d
where d.refer_code=a.refer_code 
and a.pat_id=b.pat_id 
and a.pat_id=c.pat_id 
and a.cleared= '0' or a.cleared=null 
group by a.refer_code,d.doc_name,a.pat_id,b.pat_name,c.test_rate,a.commission

end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.Pro_comm_main_FLUSH    Script Date: 26/09/2001 9:06:46 AM ******/
CREATE PROCEDURE Pro_comm_main_FLUSH

@Status int,
@refer_code varchar(10),
@pat_id varchar(10)


AS

if @Status=1
--for flush in gride from commission_main Table
begin
	select * from commission_main where refer_code=@refer_code and pat_id=@pat_id
end

if @Status=2

begin
	select s_name from test_info_sub where m_code=@refer_code and s_code=@pat_id
end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Pro_commission_flush    Script Date: 26/09/2001 9:06:46 AM ******/
CREATE PROCEDURE Pro_commission_flush

@Status int,
@pat_id varchar(10)

AS

if @Status=1
--for flush in from Commission_Details Table
begin
	select pat_id,commission from Commission_Details where pat_id=@pat_id
end

if @Status=2
--for flush in from Commission_Details Table
begin
	select m_name from test_info_main where m_code=@pat_id
end


if @Status=3
--for flush in from Commission_Details Table
begin
	select doc_name from doctor_info where refer_code=@pat_id
end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.Pro_flush_unique_id    Script Date: 26/09/2001 9:06:47 AM ******/
CREATE PROCEDURE Pro_flush_unique_id

@Status int,
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3)

AS

if @Status=1
--for flush 
begin
--	select unique_id from pat_info_sub1 where pat_id=@pat_id, m_code=@m_code, s_code=@s_code
             select unique_id from pat_info_sub1 where pat_id=@pat_id and m_code=@m_code and s_code=@s_code

end
/*
if @Status=2
begin
	select unique_id from pat_info_sub1 where pat_id=@pat_id
end
*/







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE Pro_vat_make_serial
AS

declare @Count as int
declare @var1 as int
declare @c as char(1)

DECLARE abc CURSOR FOR
SELECT pat_id FROM pat_info_main_vat 
OPEN abc

set @Count=1
set @c=1
WHILE @c = 1
begin
	FETCH NEXT FROM abc into @var1
	if @@FETCH_STATUS = 0 
	begin	
		update pat_info_main_vat set pat_id=@Count where pat_id=@var1
		set @count=@count+1

	end

	else set @c = 0
end
CLOSE abc
DEALLOCATE abc



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Report_All_Delete    Script Date: 26/09/2001 9:06:50 AM ******/
CREATE PROCEDURE Report_All_Delete

@Status int,
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3),
@filed4 varchar(500),
@type varchar(2)

 AS


if @Status=1

begin
	delete from Report_All where pat_id=@pat_id and m_code=@m_code and s_code=@s_code and field4=@filed4 and type=@type
end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Report_All_Delete1    Script Date: 26/09/2001 9:06:50 AM ******/
CREATE PROCEDURE Report_All_Delete1

@Status int,
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3),
@type varchar(2)

 AS


if @Status=1

begin
	delete from Report_All where pat_id=@pat_id and m_code=@m_code and s_code=@s_code and type=@type
end



























GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.Report_All_Delete2    Script Date: 26/09/2001 9:06:50 AM ******/
CREATE    PROCEDURE Report_All_Delete2

@Status int,
@pat_id int,
@m_code varchar(5),
@s_code varchar(5)

 AS


if @Status=1

begin
	delete from Report_All where pat_id=@pat_id and m_code=@m_code and s_code=@s_code
end

--if @Status=2

--begin
	delete from Report_All where pat_id=@pat_id and m_code=@m_code
--end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.Report_All_SELECT    Script Date: 26/09/2001 9:06:51 AM ******/
CREATE PROCEDURE Report_All_SELECT

@Status int,
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3),
@type varchar(2)

 AS


if @Status=1

begin
	select*from Report_All where pat_id=@pat_id and m_code=@m_code and s_code=@s_code and type=@type
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Report_All_SELECT3    Script Date: 20/01/2002 10:26:59 AM ******/
/****** Object:  Stored Procedure dbo.Report_All_SELECT3    Script Date: 26/09/2001 9:06:51 AM ******/
CREATE PROCEDURE Report_All_SELECT3

@Status int,
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3)

 AS


if @Status=1

begin
	select*from Report_All where pat_id=@pat_id and m_code=@m_code and s_code=@s_code
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.Report_All_Select1    Script Date: 26/09/2001 9:06:51 AM ******/
CREATE PROCEDURE Report_All_Select1

@Status int,
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3),
@type varchar(2) 

 AS


if @Status=1

begin
	select top 1*from report_all where pat_id=@pat_id and m_code=@m_code and s_code=@s_code  and type=@type
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.Report_All_Select2    Script Date: 20/01/2002 10:26:59 AM ******/
/****** Object:  Stored Procedure dbo.Report_All_Select2    Script Date: 26/09/2001 9:06:51 AM ******/
CREATE PROCEDURE Report_All_Select2

@Status int,
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3)


 AS


if @Status=1

begin
	select top 1*from report_all where pat_id=@pat_id and m_code=@m_code and s_code=@s_code
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE Report_All_Select4

@Status int,
@pat_id varchar(10),
@m_code varchar(2)


 AS


if @Status=1

begin
	select top 1*from report_all where pat_id=@pat_id and m_code=@m_code
end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  PROCEDURE Rpr_Doc_Pay3

@mode int,
@refer_code varchar(20)
As

if @mode=1
set nocount on
begin

-----9----from commission_per----------
--drop table #commission_per

select
isnull(SUM(case type when 'PATH' then comm_per else 0 end),0) as PATH_PER,
isnull(SUM(case type when 'SPATH' then comm_per else 0 end),0) as SPATH_PER,
isnull(SUM(case type when 'HISTO' then comm_per else 0 end),0) as HISTO_PER,
isnull(SUM(case type when 'X-RAY' then comm_per else 0 end),0) as [X-RAY_PER],
isnull(SUM(case type when 'ECG' then comm_per else 0 end),0) as ECG_PER,
isnull(SUM(case type when 'USG' then comm_per else 0 end),0) as USG_PER,
isnull(SUM(case type when 'ECHO' then comm_per else 0 end),0) as ECHO_PER,
isnull(SUM(case type when 'ENDO' then comm_per else 0 end),0) as ENDO_PER,
isnull(SUM(case type when 'DOPL' then comm_per else 0 end),0) as DOPL_PER,
refer_code into #Commission_Per from Commission_Per where refer_code=@refer_code group by refer_code
--select*from #commission_per
------end 9----

---10--
--drop table #group
select distinct d.dt,d.pat_id,d.pat_name,PATH=sum(d.PATH),SPATH=sum(d.SPATH),
HISTO=sum(d.HISTO),[X-RAY]=sum(d.[X-RAY]),ECG=sum(d.ECG),USG=sum(d.USG),
ECHO=sum(d.ECHO),ENDO=sum(d.ENDO),DOPL=sum(d.DOPL),d.type,comm_per=sum(d.comm_per),
VAT_amt=sum(d.VAT_amt),disc=sum(d.disc),coll_fee=sum(d.coll_fee),
adv_coll=sum(d.adv_coll),due_coll=sum(d.due_coll),test_rate=sum(d.test_rate),
d.refer_code,d.doc_name,d.doc_addr,d.doc_phone,d.doc_fax,
p.PATH_PER,p.SPATH_PER,p.HISTO_PER,p.[X-RAY_PER],p.ECG_PER,p.USG_PER,p.ECHO_PER,
p.ENDO_PER,p.DOPL_PER
into #GROUP from doc_pay d,#Commission_Per p 
group by d.dt,d.pat_id,d.pat_name,d.type,d.refer_code,d.doc_name,d.doc_addr,
d.doc_phone,d.doc_fax,p.PATH_PER,p.SPATH_PER,p.HISTO_PER,p.[X-RAY_PER],
p.ECG_PER,p.USG_PER,p.ECHO_PER,p.ENDO_PER,p.DOPL_PER
---end 10------------

--drop table #final

CREATE TABLE [#final] (
	[Sl_No] [int] IDENTITY (1, 1) NOT NULL ,
	[dt] [datetime] NULL ,
	[pat_id] [int] NULL ,
	[pat_name] [varchar] (45) NULL ,
	[PATH] [money] NULL ,
	[SPATH] [money] NULL ,
	[HISTO] [money] NULL ,
	[X-RAY] [money] NULL ,
	[ECG] [money] NULL ,
	[USG] [money] NULL ,
	[ECHO] [money] NULL ,
	[ENDO] [money] NULL ,
	[DOPL] [money] NULL ,
	[type] [varchar] (10) NULL ,
	[comm_per] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[test_rate] [money] NULL ,
	[refer_code] [varchar] (10) NULL ,
	[doc_name] [varchar] (60) NULL ,
	[doc_addr] [varchar] (60) NULL ,
	[doc_phone] [varchar] (60) NULL ,
	[doc_fax] [varchar] (60) NULL,
	[PATH_PER] [money] NULL ,
	[SPATH_PER] [money] NULL ,
	[HISTO_PER] [money] NULL ,
	[X-RAY_PER] [money] NULL ,
	[ECG_PER] [money] NULL ,
	[USG_PER] [money] NULL ,
	[ECHO_PER] [money] NULL ,
	[ENDO_PER] [money] NULL ,
	[DOPL_PER] [money] NULL ,
) ON [PRIMARY]


--drop table #final1
insert into #final select*from #group  where (PATH+SPATH+HISTO++[X-RAY]+ECG+USG+ECHO+ENDO+DOPL)<>0

select Sl_No,dt,pat_id,pat_name,Total=(PATH+SPATH+HISTO+[X-RAY]+ECG+USG+ECHO+ENDO+DOPL),PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,DOPL,
type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,doc_name,
doc_addr,doc_phone,doc_fax,PATH_PER,SPATH_PER,HISTO_PER,[X-RAY_PER],ECG_PER,USG_PER,
ECHO_PER,ENDO_PER,DOPL_PER,DUE=(test_rate+VAT_amt+coll_fee-disc-adv_coll-due_coll),
Avg_Per=((disc*100)/(PATH+SPATH+HISTO+[X-RAY]+ECG+USG+ECHO+ENDO+DOPL))
into #final1 from #final
---select*from #final1 ORDER BY pat_id

----#final2--------------
--drop table #final2
--select*from #final2
select Sl_No,dt,pat_id,pat_name,
PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,DOPL,
type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,doc_name,
doc_addr,doc_phone,doc_fax,PATH_PER,SPATH_PER,HISTO_PER,[X-RAY_PER],ECG_PER,USG_PER,
ECHO_PER,ENDO_PER,DOPL_PER,DUE=(test_rate+VAT_amt-disc-adv_coll-due_coll),
Ori_PATH=((PATH*0.01*PATH_PER)-(PATH*0.01*Avg_Per))/(0.01*PATH_PER),Ori_SPATH=((SPATH*0.01*SPATH_PER)-(SPATH*0.01*Avg_Per))/(0.01*SPATH_PER),
Ori_HISTO=((HISTO*0.01*HISTO_PER)-(HISTO*0.01*Avg_Per))/(0.01*HISTO_PER),Ori_X_RAY=(([X-RAY]*0.01*[X-RAY_PER])-([X-RAY]*0.01*Avg_Per))/(0.01*[X-RAY_PER]),
Ori_ECG=((ECG*0.01*ECG_PER)-(ECG*0.01*Avg_Per))/(0.01*ECG_PER),Ori_USG=((USG*0.01*USG_PER)-(USG*0.01*Avg_Per))/(0.01*USG_PER),
Ori_ECHO=((ECHO*0.01*ECHO_PER)-(ECHO*0.01*Avg_Per))/(0.01*ECHO_PER),Ori_ENDO=((ENDO*0.01*ENDO_PER)-(ENDO*0.01*Avg_Per))/(0.01*ENDO_PER),
Ori_DOPL=((DOPL*0.01*DOPL_PER)-(DOPL*0.01*Avg_Per))/(0.01*DOPL_PER),Avg_Per into #final2 from #final1 WHERE DUE=0

update #final2 set 
Ori_PATH=0
where Ori_PATH < 0 

update #final2 set 
Ori_SPATH=0
where Ori_SPATH < 0 

update #final2 set 
Ori_HISTO=0
where Ori_HISTO < 0
update #final2 set 
Ori_X_RAY = 0
where Ori_X_RAY < 0

update #final2 set 
Ori_ECG = 0
where Ori_ECG < 0

update #final2 set 
Ori_USG = 0
where Ori_USG < 0

update #final2 set
Ori_ECHO = 0  
where Ori_ECHO < 0

update #final2 set 
Ori_ENDO = 0
where Ori_ENDO < 0

update #final2 set 
Ori_DOPL = 0
where Ori_DOPL < 0

select f.Sl_No,f.dt,f.pat_id,f.pat_name,Total=(f.Ori_PATH+f.Ori_SPATH+f.Ori_HISTO+f.Ori_X_RAY+f.Ori_ECG+f.Ori_USG+f.Ori_ECHO+f.Ori_ENDO+f.Ori_DOPL),f.PATH,f.SPATH,f.HISTO,f.[X-RAY],f.ECG,f.USG,f.ECHO,f.ENDO,f.DOPL,f.type,f.comm_per,f.VAT_amt,f.disc,f.coll_fee,
f.adv_coll,f.due_coll,f.test_rate,f.refer_code,f.doc_name,f.doc_addr,f.doc_phone,f.doc_fax,f.PATH_PER,f.SPATH_PER,f.HISTO_PER,f.[X-RAY_PER],f.ECG_PER,f.USG_PER,f.ECHO_PER,f.ENDO_PER,f.DOPL_PER,f.DUE,Ori_PATH,f.Ori_SPATH,f.Ori_HISTO,f.Ori_X_RAY,f.Ori_ECG,f.Ori_USG,
f.Ori_ECHO,f.Ori_ENDO,f.Ori_DOPL,f.Avg_Per,a.pat_id1 into #final3 from #final2 f,pat_info_main a where a.pat_id=f.pat_id

update #final3 set pat_id1=pat_id where pat_id1=''

select*from #final3 order by sl_no

end

set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE         PROCEDURE Rpt_Booth

--@uid varchar(10),
@booth varchar (2),
@sdt datetime,
@edt datetime


AS
set nocount on

--drop table #tmp2
--drop table #tmp3
--drop table #tmp4
--drop table #tmp5
--drop table #tmp6
--drop table #tmp31
/*
--->>>---advanc>>>
select a.pat_id,adv=sum(a.adv),collect_fee=sum(a.collect_fee) 
into #tmp31 
from pat_info_sub2 a
where a.dt2 between @sdt and @edt
group by a.pat_id

select t.pat_id,a.pat_name,adv=sum(t.adv),
collect_fee=sum(t.collect_fee),vat_amt=sum(a.vat_amt),
doc_name=isnull((select doc_name from doctor_info c where
c.refer_code=a.refer_code) ,(select doc_name from doctor_info_new d where
d.pat_id=a.pat_id))
into #tmp3 from #tmp31 t,pat_info_main a
where t.pat_id=a.pat_id and a.booth=@booth
group by t.pat_id,a.pat_name,a.refer_code,a.pat_id

--select adv=sum(adv) from #tmp3 where booth='2'
--select * from #tmp3 where booth='2'

--<<<<<end advance---<<

---->>>test Rate-->>
select a.pat_id,test_rate=sum(a.test_rate) 
into #tmp2 from pat_info_sub1 a, #tmp3 t
where a.pat_id=t.pat_id group by a.pat_id
--<<<---end Test Rate --<<

select b.pat_id,disc=sum(a.disc) into #tmp4 
from pat_info_sub3 a,#tmp3 b
where a.pat_id=b.pat_id group by b.pat_id


select c.pat_id,c.pat_name,c.doc_name,c.vat_amt,
b.test_rate,c.adv,c.collect_fee,d.disc
into #tmp5 from 
#tmp2 b,#tmp3 c,#tmp4 d
where 
c.pat_id=d.pat_id
and b.pat_id=d.pat_id


select pat_id,pat_name,doc_name,
tot_bill=(vat_amt+test_rate+collect_fee-disc),adv
into #tmp6 from #tmp5
select pat_id,pat_name,doc_name,tot_bill,adv,due=(tot_bill-adv) from #tmp6

*/


--->>>---advanc>>>
select a.pat_id,adv=sum(a.adv),collect_fee=sum(a.collect_fee) 
into #tmp31 
from pat_info_sub2 a
where a.dt2 between @sdt and @edt
group by a.pat_id



select t.pat_id,a.pat_name,adv=sum(t.adv),
collect_fee=sum(t.collect_fee),vat_amt=sum(a.vat_amt),
doc_name=isnull((select doc_name from doctor_info c where
c.refer_code=a.refer_code) ,(select doc_name from doctor_info_new d where
d.pat_id=a.pat_id)),a.pat_id1
into #tmp3 from #tmp31 t,pat_info_main a
where t.pat_id=a.pat_id and a.booth=@booth
group by t.pat_id,a.pat_name,a.refer_code,a.pat_id,a.pat_id1

update #tmp3 set pat_id1=pat_id where pat_id1=''


--select adv=sum(adv) from #tmp3 where booth='2'
--select * from #tmp3 where booth='2'

--<<<<<end advance---<<

---->>>test Rate-->>
select a.pat_id,test_rate=sum(a.test_rate),t.pat_id1
into #tmp2 from pat_info_sub1 a, #tmp3 t
where a.pat_id=t.pat_id group by a.pat_id,t.pat_id1
--<<<---end Test Rate --<<

select b.pat_id,disc=sum(a.disc),b.pat_id1 into #tmp4 
from pat_info_sub3 a,#tmp3 b
where a.pat_id=b.pat_id group by b.pat_id,b.pat_id1


select c.pat_id,c.pat_name,c.doc_name,c.vat_amt,
b.test_rate,c.adv,c.collect_fee,d.disc,c.pat_id1
into #tmp5 from 
#tmp2 b,#tmp3 c,#tmp4 d
where 
c.pat_id=d.pat_id
and b.pat_id=d.pat_id


select pat_id,pat_name,doc_name,
tot_bill=(vat_amt+test_rate+collect_fee-disc),adv,pat_id1
into #tmp6 from #tmp5
select pat_id,pat_name,doc_name,tot_bill,adv,due=(tot_bill-adv),pat_id1 from #tmp6



set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE Rpt_Booth1

@uid varchar(10),
@booth varchar (2),
@sdt datetime,
@edt datetime


AS
set nocount on


--drop table #Tmp1
select distinct b.pat_id,b.test_rate into #Tmp1 from pat_info_sub1 b,pat_info_sub2 c
where b.pat_id=c.pat_id
and c.dt between @sdt and @edt 
--select*from #tmp1
--drop table #tmp2
select pat_id,Test_rate=sum(test_rate) into #tmp2 from #Tmp1 group by pat_id
--SELECT*FROM #TMP2
--end part 0--

----part1
--drop table #pat_info_sub2
select c.pat_id,c.uid,c.adv Advance,c.collect_fee,c.dt 
into #pat_info_sub2
from pat_info_sub2 c
where c.uid=@uid and c.dt between @sdt and @edt
--select*from #pat_info_sub2
---end patr1----
---part 2--
--drop table #final
select a.pat_id,a.Pat_name,a.vat_amt,a.booth,a.dt1,p.uid,Advance=sum(p.Advance),
collect_fee=sum(p.collect_fee),m.u_name Usr_Name,t.Test_rate,d.disc,
doc_name=isnull((select doc_name from doctor_info c where
c.refer_code=a.refer_code) ,(select doc_name from doctor_info_new d where
d.pat_id=a.pat_id))
--f.doc_name
into #final from pat_info_main a,#pat_info_sub2 p,
micropass m,#tmp2 t,pat_info_sub3 d
--,doctor_info f
where a.pat_id=p.pat_id
and a.pat_id=t.pat_id
and a.pat_id=d.pat_id
--and f.refer_code=a.refer_code
and m.u_id=@uid
and a.booth=@booth
group by a.pat_id,a.Pat_name,a.vat_amt,a.booth,a.dt1,
p.uid,m.u_name,t.Test_rate,d.disc,a.refer_code --f.doc_name
--select*from #final
---end part 2--


select pat_id,Pat_name,doc_name,Payable=(Test_rate+vat_amt+collect_fee-disc),Advance Collected_money,
Due=(Test_rate+vat_amt+collect_fee-disc-Advance),
Test_rate,vat_amt,collect_fee,
disc,
booth,uid,
Usr_Name,dt1 from #final


set nocount off



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Rpt_Cancer    Script Date: 26/09/2001 9:06:51 AM ******/
CREATE PROCEDURE Rpt_Cancer

@status int,
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3)

AS
if @status=1

begin
select * from report_all where pat_id=@pat_id and m_code=@m_code and s_code=@s_code

end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Rpt_Doc_Pay    Script Date: 20/01/2002 10:26:59 AM ******/
CREATE PROCEDURE Rpt_Doc_Pay

@refer_code varchar(10),
@StDate datetime,
@EdDate datetime

AS

set nocount on
CREATE TABLE [#Doc_Pay] (
	[dt] [datetime] NULL ,
	[pat_id] [int] NULL ,
	[pat_name] [varchar](60) NULL ,
	[PATH] [money] NULL ,
	[SPATH] [money] NULL ,
	[HISTO] [money] NULL ,
	[X-RAY] [money] NULL ,
	[ECG] [money] NULL ,
	[USG] [money] NULL ,
	[ECHO] [money] NULL ,
	[ENDO] [money] NULL ,
	[DOPL] [money] NULL ,
	[type] [varchar] (10) NULL ,
	[comm_per] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[test_rate] [money] NULL ,
	[refer_code] [varchar] (10) NULL ,
	[doc_name] [varchar] (60) NULL ,
	[doc_addr] [varchar] (60) NULL ,
	[doc_phone] [varchar] (60) NULL ,
	[doc_fax] [varchar] (60) NULL 

) ON [PRIMARY]
-----1------form pat_info_sub2----
--drop table #pat_info_sub2
select distinct pat_id,tmp_dt into #pat_info_sub2 from pat_info_sub2
where dt between @StDate and @EdDate

-----end 1---------------------------
-----2---
--drop table #tmp
select b.tmp_dt,b.pat_id,a.pat_name,
isnull(sum(case b.type when 'PATH' then b.test_rate else 0 end),0) as PATH_RATE,
isnull(sum(case b.type when 'SPATH' then b.test_rate else 0 end),0) as SPATH_RATE,
isnull(sum(case b.type when 'HISTO' then b.test_rate else 0 end),0) as HISTO_RATE,
isnull(sum(case b.type when 'X-RAY' then b.test_rate else 0 end),0) as [X-RAY_RATE],
isnull(sum(case b.type when 'ECG' then b.test_rate else 0 end),0) as ECG_RATE,
isnull(sum(case b.type when 'USG' then b.test_rate else 0 end),0) as USG_RATE,
isnull(sum(case b.type when 'ECHO' then b.test_rate else 0 end),0) as ECHO_RATE,
isnull(sum(case b.type when 'ENDO' then b.test_rate else 0 end),0) as ENDO_RATE,
isnull(sum(case b.type when 'DOPL' then b.test_rate else 0 end),0) as DOPL_RATE,
b.type into #tmp from pat_info_sub1 b,pat_info_main a,#pat_info_sub2 f
where a.pat_id=b.pat_id and a.pat_id=f.pat_id and a.refer_code=@refer_code
--and b.dt1 between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by b.tmp_dt,b.pat_id,a.pat_name,b.type


insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,doc_name,doc_addr,
doc_phone,doc_fax) 
select tmp_dt,pat_id,pat_name,PATH_RATE=sum(PATH_RATE),SPATH_RATE=sum(SPATH_RATE),
HISTO_RATE=sum(HISTO_RATE),[X-RAY_RATE]=sum([X-RAY_RATE]),ECG_RATE=sum(ECG_RATE),
USG_RATE=sum(USG_RATE),ECHO_RATE=sum(ECHO_RATE),ENDO_RATE=sum(ENDO_RATE),
DOPL_RATE=sum(DOPL_RATE),'',0,0,0,0,0,0,0,'','','','','' from #tmp group by tmp_dt,pat_id,pat_name
---------------end 2----

------3--test_rate----
insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,doc_name,doc_addr,
doc_phone,doc_fax) 
select f.tmp_dt,b.pat_id,'',0,0,0,0,0,0,0,0,0,'',0,0,0,0,0,0,test_rate=sum(b.test_rate),
a.refer_code,e.doc_name,e.addr,e.phone,e.fax
from pat_info_sub1 b,pat_info_main a,doctor_info e,#pat_info_sub2 f
where a.refer_code=@refer_code and a.refer_code=e.refer_code 
and a.pat_id=b.pat_id 
and a.pat_id=f.pat_id 
--and a.dt1 between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by b.pat_id,f.tmp_dt,a.refer_code,e.doc_name,e.addr,e.phone,e.fax

---end-3-----------------
--->>>--4-->>>>>>>>>for refer_code,doc_name,vat_amt,disc ------

insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 

select f.tmp_dt,a.pat_id,'',0,0,0,0,0,0,0,0,0,'',0,a.vat_amt,d.disc,0,0,0,0,
a.refer_code,e.doc_name,'','',''
from pat_info_main a,doctor_info e,pat_info_sub3 d,pat_info_sub2 c,#pat_info_sub2 f
where a.refer_code=e.refer_code and a.refer_code=@refer_code
and a.pat_id=d.pat_id and f.pat_id=a.pat_id and a.pat_id=c.pat_id
--and a.dt1 between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by f.tmp_dt,a.pat_id,a.refer_code,e.doc_name,a.vat_amt,d.disc

--->>>end 4 >>>>>>>>>>>>
-- 5 ---adv----------------------------------

select distinct f.tmp_dt,c.pat_id,a.refer_code,adv_coll=sum(c.adv)
into #adv
from pat_info_sub2 c,pat_info_main a,#pat_info_sub2 f
where a.pat_id=c.pat_id and a.pat_id=f.pat_id and c.type='adv' 
and a.refer_code=@refer_code
--and c.dt between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by a.refer_code,c.pat_id,c.type,f.tmp_dt


insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,'',0,0,0,0,0,0,0,0,0,'',0,0,0,0,adv_coll=sum(adv_coll),0,0,refer_code,'','','',''
from #adv group by tmp_dt,pat_id,refer_code

-----end 5------

----6--due-----------------------------------

select distinct f.tmp_dt,c.pat_id,a.refer_code,due_coll=sum(c.adv)
into #due
from pat_info_sub2 c,pat_info_main a,#pat_info_sub2 f
where a.pat_id=c.pat_id and a.pat_id=f.pat_id and c.type='due' 
and a.refer_code=@refer_code
--and c.dt between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by a.refer_code,c.pat_id,c.type,f.tmp_dt

insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,'',0,0,0,0,0,0,0,0,0,'',0,0,0,0,0,due_coll=sum(due_coll),0,refer_code,'','','',''
from #due group by tmp_dt,pat_id,refer_code

---end 6 ---------------------------------

---7 -collect_fee-----------------------------------

--drop table #collect_fee
select distinct a.tmp_dt,a.pat_id,c.collect_fee,a.refer_code into #collect_fee
from pat_info_sub2 c,pat_info_main a,#pat_info_sub2 f where a.pat_id=f.pat_id 
and a.refer_code=@refer_code

insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,'',0,0,0,0,0,0,0,0,0,'',0,0,0,
collect_fee=sum(collect_fee),0,0,0,refer_code,'','','','' from #collect_fee
group by tmp_dt,pat_id,refer_code

---end-7--------------------------------
----8-------

----end 8----

-----9----from commission_per----------
--drop table #commission_per
select
isnull(SUM(case type when 'PATH' then comm_per else 0 end),0) as PATH_PER,
isnull(SUM(case type when 'SPATH' then comm_per else 0 end),0) as SPATH_PER,
isnull(SUM(case type when 'HISTO' then comm_per else 0 end),0) as HISTO_PER,
isnull(SUM(case type when 'X-RAY' then comm_per else 0 end),0) as [X-RAY_PER],
isnull(SUM(case type when 'ECG' then comm_per else 0 end),0) as ECG_PER,
isnull(SUM(case type when 'USG' then comm_per else 0 end),0) as USG_PER,
isnull(SUM(case type when 'ECHO' then comm_per else 0 end),0) as ECHO_PER,
isnull(SUM(case type when 'ENDO' then comm_per else 0 end),0) as ENDO_PER,
isnull(SUM(case type when 'DOPL' then comm_per else 0 end),0) as DOPL_PER,
refer_code
into #Commission_Per
from Commission_Per where refer_code=@refer_code group by refer_code

------end 9----


----testing-----
select distinct d.dt,d.pat_id,d.pat_name,PATH=sum(d.PATH),SPATH=sum(d.SPATH),HISTO=sum(d.HISTO),
[X-RAY]=sum(d.[X-RAY]),ECG=sum(d.ECG),USG=sum(d.USG),ECHO=sum(d.ECHO),ENDO=sum(d.ENDO),
DOPL=sum(d.DOPL),d.type,comm_per=sum(d.comm_per),VAT_amt=sum(d.VAT_amt),disc=sum(d.disc),
coll_fee=sum(d.coll_fee),adv_coll=sum(d.adv_coll),due_coll=sum(d.due_coll),
test_rate=sum(d.test_rate),d.refer_code,d.doc_name,d.doc_addr,d.doc_phone,d.doc_fax,
p.PATH_PER,p.SPATH_PER,p.HISTO_PER,p.[X-RAY_PER],p.ECG_PER,p.USG_PER,p.ECHO_PER,
p.ENDO_PER,p.DOPL_PER
--DUE=(d.test_rate+d.VAT_amt+d.coll_fee-d.disc-d.adv_coll-d.due_coll)
from #doc_pay d,#Commission_Per p 
--where d.refer_code=p.refer_code
group by d.dt,d.pat_id,d.pat_name,d.type,d.refer_code,d.doc_name,d.doc_addr,
d.doc_phone,d.doc_fax,p.PATH_PER,p.SPATH_PER,p.HISTO_PER,p.[X-RAY_PER],
p.ECG_PER,p.USG_PER,p.ECHO_PER,p.ENDO_PER,p.DOPL_PER
--d.test_rate,d.VAT_amt,d.coll_fee,d.disc,d.adv_coll,d.due_coll
set nocount off









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO









----/****** Object:  Stored Procedure dbo.Rpt_Doc_Pay1    Script Date: 03/04/2002 4:10:15 PM ******
CREATE     PROCEDURE Rpt_Doc_Pay1

@refer_code varchar(10),
@StDate datetime,
@EdDate datetime

AS

set nocount on


--/*


CREATE TABLE [#Doc_Pay] (
	[dt] [datetime] NULL ,
	[pat_id] [int] NULL ,
	[pat_name] [varchar](60) NULL ,
	[PATH] [money] NULL ,
	[SPATH] [money] NULL ,
	[HISTO] [money] NULL ,
	[X-RAY] [money] NULL ,
	[ECG] [money] NULL ,
	[USG] [money] NULL ,
	[ECHO] [money] NULL ,
	[ENDO] [money] NULL ,
	[DOPL] [money] NULL ,
	[type] [varchar] (10) NULL ,
	[comm_per] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[test_rate] [money] NULL ,
	[refer_code] [varchar] (10) NULL ,
	[doc_name] [varchar] (60) NULL ,
	[doc_addr] [varchar] (60) NULL ,
	[doc_phone] [varchar] (60) NULL ,
	[doc_fax] [varchar] (60) NULL 

) ON [PRIMARY]
-----1------form pat_info_sub2----
--drop table #pat_info_sub2

select distinct pat_id,tmp_dt into #pat_info_sub2 from pat_info_sub2
where dt between @StDate and @EdDate


-----end 1---------------------------
-----2---
--drop table #tmp

select b.tmp_dt,b.pat_id,a.pat_name,
isnull(sum(case b.type when 'PATH' then b.test_rate else 0 end),0) as PATH_RATE,
isnull(sum(case b.type when 'SPATH' then b.test_rate else 0 end),0) as SPATH_RATE,
isnull(sum(case b.type when 'HISTO' then b.test_rate else 0 end),0) as HISTO_RATE,
isnull(sum(case b.type when 'X-RAY' then b.test_rate else 0 end),0) as [X-RAY_RATE],
isnull(sum(case b.type when 'ECG' then b.test_rate else 0 end),0) as ECG_RATE,
isnull(sum(case b.type when 'USG' then b.test_rate else 0 end),0) as USG_RATE,
isnull(sum(case b.type when 'ECHO' then b.test_rate else 0 end),0) as ECHO_RATE,
isnull(sum(case b.type when 'ENDO' then b.test_rate else 0 end),0) as ENDO_RATE,
isnull(sum(case b.type when 'DOPL' then b.test_rate else 0 end),0) as DOPL_RATE,
b.type,a.refer_code,e.doc_name,e.addr,e.phone,e.fax
into #tmp from pat_info_sub1 b,pat_info_main a,#pat_info_sub2 f,doctor_info e
where a.pat_id=b.pat_id and a.pat_id=f.pat_id 
and e.refer_code=a.refer_code and a.refer_code=@refer_code
group by b.tmp_dt,b.pat_id,a.pat_name,b.type,a.refer_code,e.doc_name,e.addr,e.phone,e.fax


insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,pat_name,PATH_RATE=sum(PATH_RATE),SPATH_RATE=sum(SPATH_RATE),
HISTO_RATE=sum(HISTO_RATE),[X-RAY_RATE]=sum([X-RAY_RATE]),ECG_RATE=sum(ECG_RATE),
USG_RATE=sum(USG_RATE),ECHO_RATE=sum(ECHO_RATE),ENDO_RATE=sum(ENDO_RATE),
DOPL_RATE=sum(DOPL_RATE),'',0,0,0,0,0,0,0,refer_code,doc_name,addr,phone,fax 
from #tmp group by tmp_dt,pat_id,pat_name,refer_code,doc_name,addr,phone,fax 
---------------end 2----

------3--test_rate----
insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,doc_name,doc_addr,
doc_phone,doc_fax) 
select f.tmp_dt,b.pat_id,a.pat_name,0,0,0,0,0,0,0,0,0,'',0,0,0,0,0,0,
test_rate=sum(b.test_rate),
a.refer_code,e.doc_name,e.addr,e.phone,e.fax
from pat_info_sub1 b,pat_info_main a,doctor_info e,#pat_info_sub2 f
where a.refer_code=@refer_code and a.refer_code=e.refer_code 
and a.pat_id=b.pat_id 
and a.pat_id=f.pat_id 
--and a.dt1 between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by b.pat_id,a.pat_name,f.tmp_dt,a.refer_code,e.doc_name,e.addr,e.phone,e.fax

---end-3-----------------
--->>>--4-->>>>>>>>>for refer_code,doc_name,vat_amt,disc ------

insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax)  
select f.tmp_dt,a.pat_id,a.pat_name,0,0,0,0,0,0,0,0,0,'',0,a.vat_amt,d.disc,0,0,0,0,
a.refer_code,e.doc_name,e.addr,e.phone,e.fax
from pat_info_main a,doctor_info e,pat_info_sub3 d,pat_info_sub2 c,#pat_info_sub2 f
where a.refer_code=e.refer_code and a.refer_code=@refer_code
and a.pat_id=d.pat_id and f.pat_id=a.pat_id and a.pat_id=c.pat_id

group by f.tmp_dt,a.pat_id,a.pat_name,a.refer_code,e.doc_name,e.addr,
e.phone,e.fax,a.vat_amt,d.disc

--->>>end 4 >>>>>>>>>>>>
-- 5 ---adv----------------------------------
--drop table #adv

select distinct f.tmp_dt,c.pat_id,a.pat_name,a.refer_code,e.doc_name,e.addr,e.phone,e.fax
,adv_coll=sum(c.adv)
into #adv
from pat_info_sub2 c,pat_info_main a,doctor_info e,#pat_info_sub2 f
where a.pat_id=c.pat_id and a.pat_id=f.pat_id 
and a.refer_code=e.refer_code and c.type='adv' and a.refer_code=@refer_code
--and c.dt between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by a.refer_code,e.doc_name,e.addr,e.phone,e.fax,c.pat_id,a.pat_name,c.type,f.tmp_dt


insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,pat_name,0,0,0,0,0,0,0,0,0,'',0,0,0,0,adv_coll=sum(adv_coll),0,0,
refer_code,doc_name,addr,phone,fax
from #adv group by tmp_dt,pat_id,pat_name,refer_code,doc_name,addr,phone,fax

-----end 5------

----6--due-----------------------------------
--drop table #due

select distinct f.tmp_dt,c.pat_id,a.pat_name,a.refer_code,e.doc_name,e.addr,e.phone,e.fax,
due_coll=sum(c.adv)
into #due
from pat_info_sub2 c,pat_info_main a,doctor_info e,#pat_info_sub2 f
where a.pat_id=c.pat_id and a.pat_id=f.pat_id and a.refer_code=e.refer_code
and c.type='due' and a.refer_code=@refer_code
--and c.dt between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by a.refer_code,e.doc_name,e.addr,e.phone,e.fax,c.pat_id,a.pat_name,c.type,f.tmp_dt


insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,pat_name,0,0,0,0,0,0,0,0,0,'',0,0,0,0,0,
due_coll=sum(due_coll),0,refer_code,doc_name,addr,phone,fax
from #due group by tmp_dt,pat_id,pat_name,refer_code,doc_name,addr,phone,fax

---end 6 ---------------------------------

---7 -collect_fee-----------------------------------
--drop table #collect_fee

select distinct a.tmp_dt,a.pat_id,a.pat_name,c.collect_fee,a.refer_code,
e.doc_name,e.addr,e.phone,e.fax
into #collect_fee
from pat_info_sub2 c,pat_info_main a,#pat_info_sub2 f,doctor_info e 
where a.pat_id=f.pat_id and a.refer_code=e.refer_code
and a.refer_code=@refer_code

insert into #Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,pat_name,0,0,0,0,0,0,0,0,0,'',0,0,0,
collect_fee=sum(collect_fee),0,0,0,refer_code,doc_name,addr,phone,fax from #collect_fee
group by tmp_dt,pat_id,pat_name,refer_code,doc_name,addr,phone,fax

---end-7--------------------------------
----8-------

----end 8---- 
-----9----from commission_per----------
--drop table #commission_per

select
isnull(SUM(case type when 'PATH' then comm_per else 0 end),0) as PATH_PER,
isnull(SUM(case type when 'SPATH' then comm_per else 0 end),0) as SPATH_PER,
isnull(SUM(case type when 'HISTO' then comm_per else 0 end),0) as HISTO_PER,
isnull(SUM(case type when 'X-RAY' then comm_per else 0 end),0) as [X-RAY_PER],
isnull(SUM(case type when 'ECG' then comm_per else 0 end),0) as ECG_PER,
isnull(SUM(case type when 'USG' then comm_per else 0 end),0) as USG_PER,
isnull(SUM(case type when 'ECHO' then comm_per else 0 end),0) as ECHO_PER,
isnull(SUM(case type when 'ENDO' then comm_per else 0 end),0) as ENDO_PER,
isnull(SUM(case type when 'DOPL' then comm_per else 0 end),0) as DOPL_PER,
refer_code
into #Commission_Per
from Commission_Per where refer_code=@refer_code group by refer_code
--select*from #commission_per
------end 9----
---10--
--drop table #group
select distinct d.dt,d.pat_id,d.pat_name,PATH=sum(d.PATH),SPATH=sum(d.SPATH),
HISTO=sum(d.HISTO),[X-RAY]=sum(d.[X-RAY]),ECG=sum(d.ECG),USG=sum(d.USG),
ECHO=sum(d.ECHO),ENDO=sum(d.ENDO),DOPL=sum(d.DOPL),d.type,comm_per=sum(d.comm_per),
VAT_amt=sum(d.VAT_amt),disc=sum(d.disc),coll_fee=sum(d.coll_fee),
adv_coll=sum(d.adv_coll),due_coll=sum(d.due_coll),test_rate=sum(d.test_rate),
d.refer_code,d.doc_name,d.doc_addr,d.doc_phone,d.doc_fax,
p.PATH_PER,p.SPATH_PER,p.HISTO_PER,p.[X-RAY_PER],p.ECG_PER,p.USG_PER,p.ECHO_PER,
p.ENDO_PER,p.DOPL_PER
into #GROUP from #doc_pay d,#Commission_Per p 
group by d.dt,d.pat_id,d.pat_name,d.type,d.refer_code,d.doc_name,d.doc_addr,
d.doc_phone,d.doc_fax,p.PATH_PER,p.SPATH_PER,p.HISTO_PER,p.[X-RAY_PER],
p.ECG_PER,p.USG_PER,p.ECHO_PER,p.ENDO_PER,p.DOPL_PER
---end 10------------

--drop table #final

CREATE TABLE [#final] (
	[Sl_No] [int] IDENTITY (1, 1) NOT NULL ,
	[dt] [datetime] NULL ,
	[pat_id] [int] NULL ,
	[pat_name] [varchar] (45) NULL ,
	[PATH] [money] NULL ,
	[SPATH] [money] NULL ,
	[HISTO] [money] NULL ,
	[X-RAY] [money] NULL ,
	[ECG] [money] NULL ,
	[USG] [money] NULL ,
	[ECHO] [money] NULL ,
	[ENDO] [money] NULL ,
	[DOPL] [money] NULL ,
	[type] [varchar] (10) NULL ,
	[comm_per] [money] NULL ,
	[VAT_amt] [money] NULL ,
	[disc] [money] NULL ,
	[coll_fee] [money] NULL ,
	[adv_coll] [money] NULL ,
	[due_coll] [money] NULL ,
	[test_rate] [money] NULL ,
	[refer_code] [varchar] (10) NULL ,
	[doc_name] [varchar] (60) NULL ,
	[doc_addr] [varchar] (60) NULL ,
	[doc_phone] [varchar] (60) NULL ,
	[doc_fax] [varchar] (60) NULL,
	[PATH_PER] [money] NULL ,
	[SPATH_PER] [money] NULL ,
	[HISTO_PER] [money] NULL ,
	[X-RAY_PER] [money] NULL ,
	[ECG_PER] [money] NULL ,
	[USG_PER] [money] NULL ,
	[ECHO_PER] [money] NULL ,
	[ENDO_PER] [money] NULL ,
	[DOPL_PER] [money] NULL ,
) ON [PRIMARY]


--drop table #final1
insert into #final select*from #group 
select Sl_No,dt,pat_id,pat_name,Total=(PATH+SPATH+HISTO+[X-RAY]+ECG+USG+ECHO+ENDO+DOPL),PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,DOPL,
type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,doc_name,
doc_addr,doc_phone,doc_fax,PATH_PER,SPATH_PER,HISTO_PER,[X-RAY_PER],ECG_PER,USG_PER,
ECHO_PER,ENDO_PER,DOPL_PER,DUE=(test_rate+VAT_amt-disc-adv_coll-due_coll) 
into #final1 from #final

----#final2--------------
--drop table #final2

select Sl_No,dt,pat_id,pat_name,
PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,DOPL,
type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,doc_name,
doc_addr,doc_phone,doc_fax,PATH_PER,SPATH_PER,HISTO_PER,[X-RAY_PER],ECG_PER,USG_PER,
ECHO_PER,ENDO_PER,DOPL_PER,DUE=(test_rate+VAT_amt-disc-adv_coll-due_coll),
Ori_PATH=(PATH-(disc*PATH_PER/100)),Ori_SPATH=(SPATH-(disc*SPATH_PER/100)),
Ori_HISTO=(HISTO-(disc*HISTO_PER/100)),Ori_X_RAY=([X-RAY]-(disc*[X-RAY_PER]/100)),
Ori_ECG=(ECG-(disc*ECG_PER/100)),Ori_USG=(USG-(disc*USG_PER/100)),
Ori_ECHO=(ECHO-(disc*ECHO_PER/100)),Ori_ENDO=(ENDO-(disc*ENDO_PER/100)),
Ori_DOPL=(DOPL-(disc*DOPL_PER/100)) into #final2 from #final1
WHERE DUE=0

update #final2 set 
Ori_PATH=0
where Ori_PATH < 0 

update #final2 set 
Ori_SPATH=0
where Ori_SPATH < 0 

update #final2 set 
Ori_HISTO=0
where Ori_HISTO < 0
 update #final2 set 
Ori_X_RAY = 0
where Ori_X_RAY < 0

update #final2 set 
Ori_ECG = 0
where Ori_ECG < 0

update #final2 set 
Ori_USG = 0
where Ori_USG < 0

update #final2 set
Ori_ECHO = 0  
where Ori_ECHO < 0

update #final2 set 
Ori_ENDO = 0
where Ori_ENDO < 0

update #final2 set 
Ori_DOPL = 0
where Ori_DOPL < 0


select Sl_No,dt,pat_id,pat_name,Total=(Ori_PATH+Ori_SPATH+Ori_HISTO+Ori_X_RAY+Ori_ECG+Ori_USG+Ori_ECHO+Ori_ENDO+Ori_DOPL),PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,doc_name,doc_addr,doc_phone,doc_fax,PATH_PER,SPATH_PER,HISTO_PER,[X-RAY_PER],ECG_PER,USG_PER,ECHO_PER,ENDO_PER,DOPL_PER,DUE,Ori_PATH,Ori_SPATH,Ori_HISTO,Ori_X_RAY,Ori_ECG,Ori_USG,Ori_ECHO,Ori_ENDO,Ori_DOPL from #final2












set nocount off








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



/****** Object:  Stored Procedure dbo.Rpt_Doc_Pay_new    Script Date: 20/01/2002 10:26:59 AM ******/
CREATE PROCEDURE Rpt_Doc_Pay_new

@refer_code varchar(10),
@sdt datetime,
@edt datetime

AS

set nocount on
declare @S_disc as money
set @S_disc=0.0000

select a.refer_code,a.doc_name,a.addr,
b.dt,b.pat_id,b.pat_name,b.vat_amt,c.m_code,c.s_code,
sum(case d.type when 'PATH' then c.test_rate else 0 end) as PATH_RATE,
sum(case d.type when 'SPATH' then c.test_rate else 0 end) as SPATH_RATE,
sum(case d.type when 'HISTO' then c.test_rate else 0 end) as HISTO_RATE,
sum(case d.type when 'X-RAY' then c.test_rate else 0 end) as [X-RAY_RATE],
sum(case d.type when 'ECG' then c.test_rate else 0 end) as ECG_RATE,
sum(case d.type when 'USG' then c.test_rate else 0 end) as USG_RATE,
sum(case d.type when 'ECHO' then c.test_rate else 0 end) as ECHO_RATE,
sum(case d.type when 'DOPL' then c.test_rate else 0 end) as DOPL_RATE,
sum(case d.type when 'ENDO' then c.test_rate else 0 end) as ENDO_RATE,

(case g.type when 'PATH' then g.comm_per else 0 end) as PATH_PER,
(case g.type when 'SPATH' then g.comm_per else 0 end) as SPATH_PER,
(case g.type when 'HISTO' then g.comm_per else 0 end) as HISTO_PER,
(case g.type when 'X-RAY' then g.comm_per else 0 end) as [X-RAY_PER],
(case g.type when 'ECG' then g.comm_per else 0 end) as ECG_PER,
(case g.type when 'USG' then g.comm_per else 0 end) as USG_PER,
(case g.type when 'ECHO' then g.comm_per else 0 end) as ECHO_PER,
(case g.type when 'DOPL' then g.comm_per else 0 end) as DOPL_PER,
(case g.type when 'ENDO' then g.comm_per else 0 end) as ENDO_PER,

f.disc,S_disc=@S_disc,adv into #tmp
from doctor_info a,pat_info_main b,pat_info_sub1 c,test_info_sub d,pat_info_sub2 e,pat_info_sub3 f,commission_per g
where a.Refer_code=b.Refer_code
and b.pat_id=c.pat_id
and b.pat_id=e.pat_id
and b.pat_id=f.pat_id
and c.m_code =d.m_code and c.s_code=d.s_code
and a.Refer_code=@refer_code
and b.dt between @sdt and @edt
group by a.refer_code,a.doc_name,a.addr,b.dt,b.pat_id,
b.pat_name,b.vat_amt,c.m_code,c.s_code,c.test_rate,f.disc,g.type,g.comm_per,e.adv

if (select top 1 PATH_PER from #tmp where PATH_PER<>0)=null 
	update #tmp set PATH_PER =0
else
	update #tmp set PATH_PER =(select top 1 PATH_PER from #tmp where PATH_PER<>0)
if (select top 1 SPATH_PER from #tmp where SPATH_PER<>0)=null 
	update #tmp set SPATH_PER =0
else
	update #tmp set SPATH_PER =(select top 1 SPATH_PER from #tmp where SPATH_PER<>0)
if (select top 1 HISTO_PER from #tmp where HISTO_PER<>0)=null 
	update #tmp set HISTO_PER =0
else
	update #tmp set HISTO_PER =(select top 1 HISTO_PER from #tmp where HISTO_PER<>0)
if (select top 1 [X-RAY_PER] from #tmp where [X-RAY_PER]<>0)=null 
	update #tmp set [X-RAY_PER] =0
else
	update #tmp set [X-RAY_PER] =(select top 1 [X-RAY_PER] from #tmp where [X-RAY_PER]<>0)
if (select top 1 ECG_PER from #tmp where ECG_PER<>0)=null 
	update #tmp set ECG_PER =0
else
	update #tmp set ECG_PER =(select top 1 ECG_PER from #tmp where ECG_PER<>0)
if (select top 1 USG_PER from #tmp where USG_PER<>0)=null 
	update #tmp set USG_PER =0
else
	update #tmp set USG_PER =(select top 1 USG_PER from #tmp where USG_PER<>0)
if (select top 1 ECHO_PER from #tmp where ECHO_PER<>0)=null 
	update #tmp set ECHO_PER =0
else
	update #tmp set ECHO_PER =(select top 1 ECHO_PER from #tmp where ECHO_PER<>0) 

if (select top 1 DOPL_PER from #tmp where DOPL_PER<>0)=null 
	update #tmp set DOPL_PER =0
else
	update #tmp set ENDO_PER =(select top 1 ENDO_PER from #tmp where ENDO_PER<>0)


declare @pat_id as int
declare @c as char(1)

DECLARE abc CURSOR FOR
select pat_id from #tmp group by pat_id
OPEN abc
set @c='1'
WHILE @c = '1'
begin
    FETCH NEXT FROM abc into @pat_id
   
	if @@FETCH_STATUS = 0 
    	begin
		update #tmp set S_disc=(select distinct disc from #tmp where pat_id=@pat_id ) where pat_id=@pat_id
	end
	else set @c = '0'
end
CLOSE abc
DEALLOCATE abc

select s_disc into #tmp1 from #tmp group by s_disc
set @S_disc=(select sum(s_disc) from #tmp1)
drop table #tmp1

declare @S_PATH_RATE as money select pat_id,[PATH_RATE] into #tmp2 from #tmp group by pat_id,[PATH_RATE] select @S_PATH_RATE=sum([PATH_RATE]) from #tmp2 drop table #tmp2
declare @S_SPATH_RATE as money select pat_id,[SPATH_RATE] into #tmp3 from #tmp group by pat_id,[SPATH_RATE] select @S_SPATH_RATE=sum([SPATH_RATE]) from #tmp3 drop table #tmp3
declare @S_HISTO_RATE as money select pat_id,[HISTO_RATE] into #tmp4 from #tmp group by pat_id,[HISTO_RATE] select @S_HISTO_RATE=sum([HISTO_RATE]) from #tmp4 drop table #tmp4
declare @S_XRAY_RATE as money select pat_id,[X-RAY_RATE] into #tmp5 from #tmp group by pat_id,[X-RAY_RATE] select @S_XRAY_RATE=sum([X-RAY_RATE]) from #tmp5 drop table #tmp5
declare @S_ECG_RATE as money select pat_id,[ECG_RATE] into #tmp6 from #tmp group by pat_id,[ECG_RATE] select @S_ECG_RATE=sum([ECG_RATE]) from #tmp6 drop table #tmp6
declare @S_USG_RATE as money select pat_id,[USG_RATE] into #tmp7 from #tmp group by pat_id,[USG_RATE] select @S_USG_RATE=sum([USG_RATE]) from #tmp7 drop table #tmp7
declare @S_ECHO_RATE as money select pat_id,[ECHO_RATE] into #tmp8 from #tmp group by pat_id,[ECHO_RATE] select @S_ECHO_RATE=sum([ECHO_RATE]) from #tmp8 drop table #tmp8
declare @S_DOPL_RATE as money select pat_id,[DOPL_RATE] into #tmp9 from #tmp group by pat_id,[DOPL_RATE] select @S_DOPL_RATE=sum([DOPL_RATE]) from #tmp9 drop table #tmp9
declare @S_ENDO_RATE as money select pat_id,[ENDO_RATE] into #tmp10 from #tmp group by pat_id,[ENDO_RATE] select @S_ENDO_RATE=sum([ENDO_RATE]) from #tmp10 drop table #tmp10


select a.refer_code,a.doc_name,a.addr,
b.dt,b.pat_id,b.pat_name,b.vat_amt,c.m_code,c.s_code,
sum(case d.type when 'PATH' then c.test_rate else 0 end) as PATH_RATE,S_PATH_RATE=@S_PATH_RATE,
sum(case d.type when 'SPATH' then c.test_rate else 0 end) as SPATH_RATE,S_SPATH_RATE=@S_SPATH_RATE,
sum(case d.type when 'HISTO' then c.test_rate else 0 end) as HISTO_RATE,S_HISTO_RATE=@S_HISTO_RATE,
sum(case d.type when 'X-RAY' then c.test_rate else 0 end) as [X-RAY_RATE],S_XRAY_RATE=@S_XRAY_RATE,
sum(case d.type when 'ECG' then c.test_rate else 0 end) as ECG_RATE,S_ECG_RATE=@S_ECG_RATE,
sum(case d.type when 'USG' then c.test_rate else 0 end) as USG_RATE,S_USG_RATE=@S_USG_RATE,
sum(case d.type when 'ECHO' then c.test_rate else 0 end) as ECHO_RATE,S_ECHO_RATE=@S_ECHO_RATE,
sum(case d.type when 'DOPL' then c.test_rate else 0 end) as DOPL_RATE,S_DOPL_RATE=@S_DOPL_RATE,
sum(case d.type when 'ENDO' then c.test_rate else 0 end) as ENDO_RATE,S_ENDO_RATE=@S_ENDO_RATE,

PATH_PER=(select top 1 PATH_PER from #tmp),
SPATH_PER=(select top 1 SPATH_PER from #tmp),
HISTO_PER=(select top 1 HISTO_PER from #tmp),
[X-RAY_PER]=(select top 1 [X-RAY_PER] from #tmp),
ECG_PER=(select top 1 ECG_PER from #tmp),
USG_PER=(select top 1 USG_PER from #tmp),
ECHO_PER=(select top 1 ECHO_PER from #tmp),
DOPL_PER=(select top 1 DOPL_PER from #tmp),
ENDO_PER=(select top 1 ENDO_PER from #tmp),

f.disc,S_disc=@S_disc,adv=( select distinct sum(e.adv) from #tmp group by pat_id)
from doctor_info a,pat_info_main b,pat_info_sub1 c,test_info_sub d,pat_info_sub2 e,pat_info_sub3 f,commission_per g
where a.Refer_code=b.Refer_code
and b.pat_id=c.pat_id
and b.pat_id=e.pat_id
and b.pat_id=f.pat_id
and c.m_code =d.m_code and c.s_code=d.s_code
and a.Refer_code=@refer_code
and b.dt between @sdt and @edt
group by a.refer_code,a.doc_name,a.addr,b.dt,b.pat_id,
b.pat_name,b.vat_amt,c.m_code,c.s_code,c.test_rate,f.disc,g.type,g.comm_per,e.adv

drop table #tmp









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



-- EXEC Rpt_Doc_pay2 1,'100','01 feb 2006','28 feb 2006'

CREATE       PROC Rpt_Doc_pay2
@mode int,
@refer_code varchar(10),
@StDate datetime,
@EdDate datetime
As

set nocount on

if @mode=1
begin

delete from doc_pay
--select*from doc_pay
select pat_id,dt=max(dt),tmp_dt into #pat_info_sub2N
from pat_info_sub2 group by pat_id,tmp_dt

select distinct pat_id,tmp_dt into #pat_info_sub2 from #pat_info_sub2N
where dt between @StDate and @EdDate

--select*from #pat_info_sub2

-----end 1---------------------------
-----2---
-- drop table #tmp
select b.tmp_dt,b.pat_id,a.pat_name,
isnull(sum(case b.type when 'PATH' then b.test_rate else 0 end),0) as PATH_RATE,
isnull(sum(case b.type when 'SPATH' then b.test_rate else 0 end),0) as SPATH_RATE,
isnull(sum(case b.type when 'HISTO' then b.test_rate else 0 end),0) as HISTO_RATE,
isnull(sum(case b.type when 'X-RAY' then b.test_rate else 0 end),0) as [X-RAY_RATE],
isnull(sum(case b.type when 'ECG' then b.test_rate else 0 end),0) as ECG_RATE,
isnull(sum(case b.type when 'USG' then b.test_rate else 0 end),0) as USG_RATE,
isnull(sum(case b.type when 'ECHO' then b.test_rate else 0 end),0) as ECHO_RATE,
isnull(sum(case b.type when 'ENDO' then b.test_rate else 0 end),0) as ENDO_RATE,
isnull(sum(case b.type when 'DOPL' then b.test_rate else 0 end),0) as DOPL_RATE,
b.type,a.refer_code,e.doc_name,e.addr,e.phone,e.fax
into #tmp from pat_info_sub1 b,pat_info_main a,#pat_info_sub2 f,doctor_info e
where b.type<>'Misc' and a.pat_id=b.pat_id and a.pat_id=f.pat_id 
and e.refer_code=a.refer_code and a.refer_code=@refer_code
group by b.tmp_dt,b.pat_id,a.pat_name,b.type,a.refer_code,e.doc_name,e.addr,e.phone,e.fax

insert into Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,pat_name,PATH_RATE=sum(PATH_RATE),SPATH_RATE=sum(SPATH_RATE),
HISTO_RATE=sum(HISTO_RATE),[X-RAY_RATE]=sum([X-RAY_RATE]),ECG_RATE=sum(ECG_RATE),
USG_RATE=sum(USG_RATE),ECHO_RATE=sum(ECHO_RATE),ENDO_RATE=sum(ENDO_RATE),
DOPL_RATE=sum(DOPL_RATE),'',0,0,0,0,0,0,0,refer_code,doc_name,addr,phone,fax 
from #tmp group by tmp_dt,pat_id,pat_name,refer_code,doc_name,addr,phone,fax 
---------------end 2----

--select*from #doc_pay
------3--test_rate----
insert into Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,doc_name,doc_addr,
doc_phone,doc_fax) 
select f.tmp_dt,b.pat_id,a.pat_name,0,0,0,0,0,0,0,0,0,'',0,0,0,0,0,0,
test_rate=sum(b.test_rate),
a.refer_code,e.doc_name,e.addr,e.phone,e.fax
from pat_info_sub1 b,pat_info_main a,doctor_info e,#pat_info_sub2 f
where b.type<>'Misc' and a.refer_code=@refer_code and a.refer_code=e.refer_code 
and a.pat_id=b.pat_id and a.pat_id=f.pat_id group by b.pat_id,a.pat_name,f.tmp_dt,a.refer_code,e.doc_name,e.addr,e.phone,e.fax

---end-3-----------------

-->>>>>discount--->

select d.pat_id,disc=sum(d.disc)
into #ttmp1 from pat_info_sub3 d,#pat_info_sub2 f
where f.pat_id=d.pat_id
group by d.pat_id
--select*from #ttmp1
--<<<<end Discount---

--->>>--4-->>>>>>>>>for refer_code,doc_name,vat_amt------

insert into Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) select f.tmp_dt,a.pat_id,a.pat_name,0,0,0,0,0,0,0,0,0,'',0,a.vat_amt,0,0,0,0,0,
a.refer_code,e.doc_name,e.addr,e.phone,e.fax
from pat_info_main a,doctor_info e,
--pat_info_sub2 c,
#pat_info_sub2 f
where a.refer_code=e.refer_code and a.refer_code=@refer_code
--and a.pat_id=d.pat_id 
and f.pat_id=a.pat_id
-- and a.pat_id=c.pat_id
--group by f.tmp_dt,a.pat_id,a.pat_name,a.refer_code,e.doc_name,e.addr,
--e.phone,e.fax,a.vat_amt
--,d.disc
---select*from #pat_info_sub2
--select pat_id,disc=sum(disc) from #doc_pay group by pat_id 
--select * from #doc_pay order by pat_id 

--->>>--new 4-->>>>>>>>>for Discount Only------

insert into Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select f.tmp_dt,a.pat_id,a.pat_name,0,0,0,0,0,0,0,0,0,'',0,0,disc=sum(d.disc),0,0,0,0,
a.refer_code,e.doc_name,e.addr,e.phone,e.fax
from pat_info_main a,#ttmp1 d,doctor_info e,#pat_info_sub2 f
where a.refer_code=e.refer_code and a.refer_code=@refer_code
and a.pat_id=d.pat_id 
and f.pat_id=a.pat_id 
--and a.pat_id=c.pat_id
group by f.tmp_dt,a.pat_id,a.pat_name,a.refer_code,e.doc_name,e.addr,
e.phone,e.fax
--,a.vat_amt
--,d.disc

--select pat_id,disc=sum(disc) from #doc_pay group by pat_id 
--select * from #doc_pay order by pat_id 
--->>>end new 4 >>>>>>>>>>>>

--->>>end 4 >>>>>>>>>>>>
-- 5 ---adv----------------------------------
--drop table #adv
--select*from #Doc_Pay

select distinct f.tmp_dt,c.pat_id,a.pat_name,a.refer_code,e.doc_name,e.addr,e.phone,e.fax
,adv_coll=sum(c.adv) into #adv from pat_info_sub2 c,pat_info_main a,doctor_info e,#pat_info_sub2 f
where a.pat_id=c.pat_id and a.pat_id=f.pat_id 
and a.refer_code=e.refer_code and c.type='adv' and a.refer_code=@refer_code
--and c.dt between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by a.refer_code,e.doc_name,e.addr,e.phone,e.fax,c.pat_id,a.pat_name,c.type,f.tmp_dt
--select*from #adv

insert into Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,pat_name,0,0,0,0,0,0,0,0,0,'',0,0,0,0,adv_coll=sum(adv_coll),0,0,
refer_code,doc_name,addr,phone,fax
from #adv group by tmp_dt,pat_id,pat_name,refer_code,doc_name,addr,phone,fax

-----end 5------

----6--due-----------------------------------
--drop table #due

select distinct f.tmp_dt,c.pat_id,a.pat_name,a.refer_code,e.doc_name,e.addr,e.phone,e.fax,
due_coll=sum(c.adv)
into #due 
from pat_info_sub2 c,pat_info_main a,doctor_info e,#pat_info_sub2 f
where a.pat_id=c.pat_id and a.pat_id=f.pat_id and a.refer_code=e.refer_code
and c.type='due' and a.refer_code=@refer_code
--and c.dt between '2002-03-12 11:54:00.000' and '2002-03-13 12:32:00.000'
group by a.refer_code,e.doc_name,e.addr,e.phone,e.fax,c.pat_id,a.pat_name,c.type,f.tmp_dt

insert into Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,pat_name,0,0,0,0,0,0,0,0,0,'',0,0,0,0,0,
due_coll=sum(due_coll),0,refer_code,doc_name,addr,phone,fax
from #due group by tmp_dt,pat_id,pat_name,refer_code,doc_name,addr,phone,fax
--select*from #due

---end 6 ---------------------------------

---7 -collect_fee-----------------------------------
--drop table #collect_fee

select distinct a.tmp_dt,a.pat_id,a.pat_name,c.collect_fee,a.refer_code,
e.doc_name,e.addr,e.phone,e.fax
into #collect_fee
from pat_info_sub2 c,pat_info_main a,#pat_info_sub2 f,doctor_info e 
where a.pat_id=f.pat_id 
and a.pat_id=c.pat_id
and a.refer_code=e.refer_code
and a.refer_code=@refer_code

insert into Doc_Pay(dt,pat_id,pat_name,PATH,SPATH,HISTO,[X-RAY],ECG,USG,ECHO,ENDO,
DOPL,type,comm_per,VAT_amt,disc,coll_fee,adv_coll,due_coll,test_rate,refer_code,
doc_name,doc_addr,doc_phone,doc_fax) 
select tmp_dt,pat_id,pat_name,0,0,0,0,0,0,0,0,0,'',0,0,0,
collect_fee=sum(collect_fee),0,0,0,refer_code,doc_name,addr,phone,fax from #collect_fee
group by tmp_dt,pat_id,pat_name,refer_code,doc_name,addr,phone,fax
--select * from #collect_fee

---end-7--------------------------------
select msg='Process Completed'
end

set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Rpt_Doctor_Info    Script Date: 26/09/2001 9:06:52 AM ******/
CREATE PROCEDURE Rpt_Doctor_Info

@Status int,
@refer_code varchar (10)

AS

if @Status=1
begin
	select * from doctor_info
end

if @Status=2
begin
	select * from doctor_info where refer_code =@refer_code
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE  PROCEDURE Rpt_Doctor_Info_New

@status char(1),
@doc_name varchar(45),
@StDate datetime,
@EdDate datetime

 AS
set nocount on
If @status='1'

begin
--	select a.pat_id,a.doc_name,a.addr,a.phone,a.fax,a.email,a.dt,b.pat_name
--	from doctor_info_new a,pat_info_main b
--	where a.pat_id=b.pat_id and a.doc_name<>'' and a.pat_id<>'' 
--	and a.dt between @StDate and @EdDate


select a.pat_id,a.doc_name,a.addr,a.phone,a.fax,a.email,a.dt,b.pat_name,b.pat_id1
into #tmp1 from doctor_info_new a,pat_info_main b
where a.pat_id=b.pat_id and a.doc_name<>'' and a.pat_id<>'' 
and a.dt between @StDate and @EdDate

update #tmp1 set pat_id1=pat_id where pat_id1=''
select*from #tmp1

end
If @status='2'

begin

--	select a.pat_id,a.doc_name,a.addr,a.phone,a.fax,a.email,a.dt,b.pat_name
--	from doctor_info_new a,pat_info_main b
--	where a.doc_name=@doc_name and a.pat_id=b.pat_id and a.doc_name<>'' and a.pat_id<>'' 
--	and a.dt between @StDate and @EdDate

	select a.pat_id,a.doc_name,a.addr,a.phone,a.fax,a.email,a.dt,b.pat_name,b.pat_id1
	into #tmp2 from doctor_info_new a,pat_info_main b
	where a.doc_name=@doc_name and a.pat_id=b.pat_id and a.doc_name<>'' and a.pat_id<>'' 
	and a.dt between @StDate and @EdDate

	update #tmp2 set pat_id1=pat_id where pat_id1=''

	select*from #tmp2


end



set nocount off



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE     PROCEDURE Rpt_Pat_info

@pat_id int

AS
set nocount on



select a.pat_id,a.m_code,a.s_code,
s_name=(select s_name from test_info_sub where a.m_code=test_info_sub.m_code and a.s_code=test_info_sub.s_code and pat_id=@pat_id),
a.test_rate,a.delv_dt,a.uid,
usr_name=(select u_name from micropass where u_id=(select distinct uid from pat_info_sub1 where pat_id=@pat_id)),
CDt_TM=(select dt from pat_info_main where pat_id=@pat_id),
Booth=(select booth from pat_info_main where pat_id=@pat_id),
pat_name=(select pat_name from pat_info_main where pat_id=@pat_id),
Sex=(select sex from pat_info_main where pat_id=@pat_id),
Age=(select age from pat_info_main where pat_id=@pat_id),
Address=(select addr from pat_info_main where pat_id=@pat_id),
Phone=(select phone from pat_info_main where pat_id=@pat_id),
vat_per=(select vat_per from pat_info_main where pat_id=@pat_id),
vat_amt=(select vat_amt from pat_info_main where pat_id=@pat_id),
Discount=(select sum(disc) from pat_info_sub3 where pat_id=@pat_id),
Doctor_Name=isnull((select doc_name from doctor_info 
where doctor_info.refer_code=(select refer_code from pat_info_main where pat_id=@pat_id)),
(select doc_name from doctor_info_new where pat_id=@pat_id)),
Advance=isnull((select sum(adv) from pat_info_sub2 where pat_id=@pat_id),0),
Coll_Fee_Sum=isnull((select sum(collect_fee) from pat_info_sub2 where pat_id=@pat_id),0),
Total_Rate=(select sum(test_rate) from pat_info_sub1 where pat_id=@pat_id),
Total_Amt=((select sum(test_rate) from pat_info_sub1 where pat_id=@pat_id)+(select vat_amt from pat_info_main where pat_id=@pat_id)+isnull((select sum(collect_fee) from pat_info_sub2 where pat_id=@pat_id),0)),
Due=(select sum(test_rate) from pat_info_sub1 where pat_id=@pat_id)
-(select sum(disc) from pat_info_sub3 where pat_id=@pat_id)
+(select sum(vat_amt) from pat_info_main where pat_id=@pat_id)
-isnull((select sum(adv) from pat_info_sub2 where pat_id=@pat_id),0)
+isnull((select sum(collect_fee) from pat_info_sub2 where pat_id=@pat_id),0),
pat_id1=(select pat_id1 from pat_info_main where pat_id=@pat_id),
refer_type=(select refer_type from pat_info_main where pat_id=@pat_id)

into #tmp1 from pat_info_sub1 a where pat_id=@pat_id

declare @pat_id1 varchar(50)
set @pat_id1=(select top 1 pat_id1 from #tmp1)
if @pat_id1=''
	begin
		update #tmp1 set pat_id1=@pat_id
	end

select*from #tmp1

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Rpt_Test_Info    Script Date: 26/09/2001 9:06:52 AM ******/
CREATE  PROCEDURE Rpt_Test_Info

@Status int,
@m_code varchar (2)


AS
set nocount on



if @Status=1
begin

CREATE TABLE #tmp1
	(m_code int null,
	m_name varchar(100) null,
	s_code int null,
	s_name varchar(100)null,
	type varchar(50) null,
	rate money null
)

insert into #tmp1
select distinct a.m_code,a.m_name,b.s_code,b.s_name,b.type,c.rate from
test_info_main a,test_info_sub b, test_info_rate c where
a.m_code=b.m_code and a.m_code=c.m_code and b.s_code=c.s_code
select*from #tmp1
end



if @Status=2
begin

CREATE TABLE #tmp2
	(m_code int null,
	m_name varchar(100) null,
	s_code int null,
	s_name varchar(100)null,
	type varchar(50) null,
	rate money null
)

insert into #tmp2
select distinct a.m_code,a.m_name,b.s_code,b.s_name,b.type,c.rate from
test_info_main a,test_info_sub b, test_info_rate c where
a.m_code=b.m_code and a.m_code=c.m_code and b.s_code=c.s_code and a.m_code=@m_code

select*from #tmp2

end

set nocount off








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE       PROCEDURE Rpt_VAT_Pat


@StDate datetime,
@EdDate datetime


AS

set nocount on	

/*
--patr 1--
--drop table #pat_info_sub1
select b.pat_id,sum(b.test_rate) Tot_test_rate into #pat_info_sub1 
from pat_info_sub1_VAT b group by b.pat_id
--select*from #pat_info_sub1
--end part-1--
--part-2
--drop table #tmp1
select s.pat_id,b.m_code,b.s_code,t.s_name,
b.delv_dt,b.test_rate,s.Tot_test_rate
into #tmp1 from pat_info_sub1_VAT b,#pat_info_sub1 s,test_info_sub t
where b.pat_id=s.pat_id and b.m_code=t.m_code and b.s_code=t.s_code
---select * from #tmp1 p
---end part-2

--part-3
--drop table #tmp2

select a.pat_id,a.pat_name,a.sex,a.age,a.refer_code,r.doc_name,
New_Doc=(select doc_name from doctor_info_new where doctor_info_new.pat_id=a.pat_id),
a.addr,a.phone,a.fax,a.email,
a.uid,m.u_name,a.vat_per,a.vat_amt,a.booth,a.dt1,d.disc into #tmp2
from pat_info_main_VAT a,pat_info_sub3_VAT d,doctor_info r,
micropass m
where a.refer_code=r.refer_code
and a.pat_id=d.pat_id and m.u_id=a.uid



--end part3

---part 4-------------------------------
--drop table #pat_info_sub2
select distinct c.pat_id,adv=sum(c.adv),collect_fee=sum(c.collect_fee) into #pat_info_sub2
from pat_info_sub2_VAT c
group by c.pat_id
--select*from #pat_info_sub2
---end part4--------------------------------

--part 5-----
--drop table #final
--drop table #final1
select p.pat_id,p.m_code,p.s_code,p.s_name,p.delv_dt,p.test_rate,p.Tot_test_rate,t.pat_name,
t.sex,t.age,t.refer_code,t.doc_name,t.New_Doc,t.addr,t.phone,t.fax,t.email,
t.uid,t.u_name,t.vat_per,t.vat_amt,t.booth,t.disc,t.dt1,f.adv,f.collect_fee
into #final from #tmp1 p,#tmp2 t,#pat_info_sub2 f

where p.pat_id=t.pat_id and f.pat_id=p.pat_id and t.pat_id=f.pat_id
select pat_id,pat_name,age,sex,addr,phone,refer_code,doc_name,New_Doc,delv_dt,
m_code,s_code,s_name,vat_per,vat_amt,test_rate,Tot_test_rate,disc,
uid,u_name,booth,dt1,adv,collect_fee,
Net_Amount=(Tot_test_rate+vat_amt+collect_fee-disc),
Advance=((Tot_test_rate+vat_amt+collect_fee-disc)-(adv)),
Due=((adv)-(Tot_test_rate+vat_amt+collect_fee-disc))
into #final1 from #final

--end part 5
--part 6-----------------------------------------------------------
--drop table #VAT
CREATE TABLE [#VAT] (
	[VAT_id] [int] IDENTITY (1, 1) NOT NULL ,
	[pat_id] [int] NULL ,
	[pat_name] [varchar] (45) NULL ,
	[age] [varchar] (17) NULL ,
	[sex] [varchar] (10) NULL ,
	[addr] [varchar] (200) NULL ,
	[phone] [varchar] (25) NULL ,
	[refer_code] [varchar] (10) NULL ,
	[doc_name] [varchar] (45) NULL ,	
	[New_Doc] [varchar] (45) NULL ,
	[delv_dt] [datetime] NULL ,
	[m_code] [varchar] (5) NULL ,
	[s_code] [varchar] (7) NULL ,
	[s_name] [varchar] (60) NULL ,
	[vat_per] [money] NULL ,
	[vat_amt] [money] NULL ,
	[test_rate] [money] NULL ,
	[Tot_test_rate] [money] NULL ,
	[disc] [money] NULL ,
	[uid] [varchar] (10) NULL ,
	[u_name] [varchar] (200) NULL ,
	[booth] [varchar] (2) NULL ,
	[dt1] [datetime] NULL ,
	[adv] [money] NULL ,
	[collect_fee] [money] NULL ,
	[Net_Amount] [money] NULL ,
	[Advance] [money] NULL ,
	[Due] [money] NULL ,
)


insert into #vat(pat_id,pat_name,age,sex,addr,phone,refer_code,doc_name,New_Doc,delv_dt,m_code,s_code,s_name,vat_per,vat_amt,test_rate,Tot_test_rate,disc,uid,u_name,booth,dt1,adv,collect_fee,Net_Amount,Advance,Due)
          select pat_id,pat_name,age,sex,addr,phone,refer_code,doc_name,New_Doc,delv_dt,m_code,s_code,s_name,vat_per,vat_amt,test_rate,Tot_test_rate,disc,uid,u_name,booth,dt1,adv,collect_fee,Net_Amount,Advance,Due from #final1

--------------UPDAT PAT_ID--------------------------------------------

declare @Count as int
declare @var1 as int
declare @c as char(1)

DECLARE abc CURSOR FOR
SELECT distinct pat_id FROM #vat 
OPEN abc

set @Count=1
set @c=1
WHILE @c = 1
begin
	FETCH NEXT FROM abc into @var1
	if @@FETCH_STATUS = 0 
	begin	
		update #vat set pat_id=@Count where pat_id=@var1
		set @count=@count+1

	end

	else set @c = 0
end
CLOSE abc
DEALLOCATE abc

---------------------------------------------------------------------------

select*from #vat where dt1 between @StDate and @EdDate order by pat_id
---end part 6------

*/


--patr 1--
--drop table #pat_info_sub1
select b.pat_id,sum(b.test_rate) Tot_test_rate into #pat_info_sub1 
from pat_info_sub1_VAT b group by b.pat_id
--select*from #pat_info_sub1
--end part-1--
--part-2
--drop table #tmp1
select s.pat_id,b.m_code,b.s_code,t.s_name,
b.delv_dt,b.test_rate,s.Tot_test_rate
into #tmp1 from pat_info_sub1_VAT b,#pat_info_sub1 s,test_info_sub t
where b.pat_id=s.pat_id and b.m_code=t.m_code and b.s_code=t.s_code
---select * from #tmp1 p
---end part-2

--part-3
--drop table #tmp2

select a.pat_id,a.pat_name,a.sex,a.age,a.refer_code,r.doc_name,
New_Doc=(select doc_name from doctor_info_new where doctor_info_new.pat_id=a.pat_id),
a.addr,a.phone,a.fax,a.email,
a.uid,m.u_name,a.vat_per,a.vat_amt,a.booth,a.dt1,d.disc,a.pat_id1 into #tmp2
from pat_info_main_VAT a,pat_info_sub3_VAT d,doctor_info r,
micropass m
where a.refer_code=r.refer_code
and a.pat_id=d.pat_id and m.u_id=a.uid



--end part3

---part 4-------------------------------
--drop table #pat_info_sub2
select distinct c.pat_id,adv=sum(c.adv),collect_fee=sum(c.collect_fee) into #pat_info_sub2
from pat_info_sub2_VAT c
group by c.pat_id
--select*from #pat_info_sub2
---end part4--------------------------------

--part 5-----
--drop table #final
--drop table #final1
select distinct p.pat_id,p.m_code,p.s_code,p.s_name,p.delv_dt,p.test_rate,p.Tot_test_rate,t.pat_name,
t.sex,t.age,t.refer_code,t.doc_name,t.New_Doc,t.addr,t.phone,t.fax,t.email,
t.uid,t.u_name,t.vat_per,t.vat_amt,t.booth,t.disc,t.dt1,f.adv,f.collect_fee,t.pat_id1
into #final from #tmp1 p,#tmp2 t,#pat_info_sub2 f
where p.pat_id=t.pat_id and f.pat_id=p.pat_id and t.pat_id=f.pat_id

select distinct pat_id,pat_id1,pat_name,age,sex,addr,phone,refer_code,doc_name,New_Doc,delv_dt,
m_code,s_code,s_name,vat_per,vat_amt,test_rate,Tot_test_rate,disc,
uid,u_name,booth,dt1,adv,collect_fee,
Net_Amount=(Tot_test_rate+vat_amt+collect_fee-disc),
Advance=((Tot_test_rate+vat_amt+collect_fee-disc)-(adv)),
Due=((adv)-(Tot_test_rate+vat_amt+collect_fee-disc))
into #final1 from #final

--end part 5
--part 6-----------------------------------------------------------
--drop table #VAT
CREATE TABLE [#VAT] (
--	[VAT_id] [int] IDENTITY (1, 1) NOT NULL ,
	[pat_id] [int] NULL ,
	[pat_name] [varchar] (45) NULL ,
	[age] [varchar] (17) NULL ,
	[sex] [varchar] (10) NULL ,
	[addr] [varchar] (200) NULL ,
	[phone] [varchar] (25) NULL ,
	[refer_code] [varchar] (10) NULL ,
	[doc_name] [varchar] (45) NULL ,	
	[New_Doc] [varchar] (45) NULL ,
	[delv_dt] [datetime] NULL ,
	[m_code] [varchar] (5) NULL ,
	[s_code] [varchar] (7) NULL ,
	[s_name] [varchar] (60) NULL ,
	[vat_per] [money] NULL ,
	[vat_amt] [money] NULL ,
	[test_rate] [money] NULL ,
	[Tot_test_rate] [money] NULL ,
	[disc] [money] NULL ,
	[uid] [varchar] (10) NULL ,
	[u_name] [varchar] (200) NULL ,
	[booth] [varchar] (2) NULL ,
	[dt1] [datetime] NULL ,
	[adv] [money] NULL ,
	[collect_fee] [money] NULL ,
	[Net_Amount] [money] NULL ,
	[Advance] [money] NULL ,
	[Due] [money] NULL ,
	[pat_id1] [varchar](50) NULL
)


insert into #vat(pat_id,pat_name,age,sex,addr,phone,refer_code,doc_name,New_Doc,delv_dt,m_code,s_code,s_name,vat_per,vat_amt,test_rate,Tot_test_rate,disc,uid,u_name,booth,dt1,adv,collect_fee,Net_Amount,Advance,Due,pat_id1)
select distinct pat_id,pat_name,age,sex,addr,phone,refer_code,doc_name,New_Doc,delv_dt,m_code,s_code,s_name,vat_per,vat_amt,test_rate,Tot_test_rate,disc,uid,u_name,booth,dt1,adv,collect_fee,Net_Amount,Advance,Due,pat_id1 from #final1

--------------UPDAT PAT_ID--------------------------------------------

declare @Count as int
declare @var1 as int
declare @c as char(1)

DECLARE abc CURSOR FOR
SELECT distinct pat_id FROM #vat 
OPEN abc

set @Count=1
set @c=1
WHILE @c = 1
begin
	FETCH NEXT FROM abc into @var1
	if @@FETCH_STATUS = 0 
	begin	
		update #vat set pat_id=@Count where pat_id=@var1
		set @count=@count+1

	end

	else set @c = 0
end
CLOSE abc
DEALLOCATE abc

------------
select distinct *from #vat where dt1 between @StDate and @EdDate order by pat_id
---end part 6------

set nocount off















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.Rpt_for_All    Script Date: 26/09/2001 9:06:52 AM ******/
CREATE PROCEDURE Rpt_for_All

@status int,
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3)

AS
if @status=1

begin
--select * from report_all where pat_id=@pat_id and m_code=@m_code and s_code=@s_code

--select a.pat_id,
--b.pat_name,b.sex,b.addr,b.phone,b.fax,b.email,b.dt,doc_name =(
--	select doc_name from doctor_info where refer_code=(
--        select refer_code from pat_info_main where pat_id=@pat_id)),

--a.m_code,a.s_code,a.field1,a.field2,a.field3,a.field3,a.field4,a.field5,a.field6,
--a.field7,a.field8,a.field9,a.field10,a.field11,a.field12,a.field13,
--a.field14,a.field15,a.dt,a.type

--from report_all a,pat_info_main b
--where b.pat_id=a.pat_id and a.pat_id=@pat_id and a.m_code=@m_code and a.s_code=@s_code

/*select a.pat_id,

a.m_code,a.s_code,a.field1,a.field2,a.field3,a.field3,a.field4,a.field5,a.field6,
a.field7,a.field8,a.field9,a.field10,a.field11,a.field12,a.field13,
a.field14,a.field15,a.dt,a.type,
b.pat_name,b.sex,b.addr,b.phone,b.fax,b.email,b.dt,doc_name =(
	select doc_name from doctor_info where refer_code=(
        select refer_code from pat_info_main where pat_id=@pat_id))

from report_all a,pat_info_main b
where b.pat_id=a.pat_id and a.pat_id=@pat_id and a.m_code=@m_code and a.s_code=@s_code*/

select a.field1,a.field2,a.field3,a.field3,a.field4,a.field5,a.field6,
a.field7,a.field8,a.field9,a.field10,a.field11,a.field12,a.field13,
a.field14,a.field15,a.dt,a.type,
b.pat_name,b.sex,b.addr,b.phone,b.fax,b.email,b.dt,doc_name =(
	select doc_name from doctor_info where refer_code=(
        select refer_code from pat_info_main where pat_id=@pat_id))

from report_all a,pat_info_main b
where b.pat_id=a.pat_id and a.pat_id=@pat_id and a.m_code=@m_code and a.s_code=@s_code



end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




/****** Object:  Stored Procedure dbo.Rt    Script Date: 29/04/2002 12:40:14 AM ******/

/****** Object:  Stored Procedure dbo.Rt    Script Date: 24/04/2002 05:16:19 PM ******/
CREATE PROCEDURE Rt

@emp_id varchar(10)
AS

set nocount on

--declare @emp_id as varchar(10)
declare @JoinDate as datetime
--set @emp_id='001'
set @Joindate=(select join_date from emp_info where emp_id=@emp_id)

--drop table #JoinLeave

select leave_type,
Join_Yr_Leave=round(Celing*
(365-datediff(day,
cast((cast(year(@JoinDate)as varchar(4)) + '-01-01') as datetime),@Joindate))
/365,0)
into #JoinLeave from lv_setup

---end ok1------------

--ok2 employee's year duration----------------------

declare @EdGate as datetime
set @EdGate=(select getdate())

declare @Yr_Duration as money
set @Yr_Duration=(select Yr_Duration=DATEDIFF (year,@Joindate,@EdGate )+1)


-----end ok2---------------
---ok 3--------------------------------------------------------------
if @Yr_Duration =1
	begin
--			
		declare @EdDate as datetime
		set @EdDate=(select getdate())

--		drop table #tmp1

		select a.Emp_ID,e.Emp_Name,e.join_date,a.Leave_st_Date,a.Leave_ed_date,
		a.Leave_Type,a.Total_Leave
		into #tmp1 from leave a,emp_info e
		where e.Emp_ID=a.Emp_ID and a.Emp_ID=@emp_id
		and e.join_date between e.join_date and @EdDate

		-----------------------------------------------
--		drop table #tmp2

		select t.Emp_ID,t.Emp_Name,t.join_date,t.Leave_Type,
		sum_leave=sum(t.Total_Leave),s.celing into #tmp2 from #tmp1 t,lv_setup s
		where t.leave_type=s.leave_type
		group by t.Emp_ID,t.Emp_Name,t.join_date,t.Leave_Type,s.celing
		

--		drop table #tmp3

		select t.Emp_ID,t.Emp_Name,t.join_date,
		t.sum_leave Tot_Availed_Leave,t.celing,j.leave_type,j.Join_Yr_Leave 
		into #tmp3 from #tmp2 t,#JoinLeave j where j.leave_type=t.Leave_Type

--		drop table #tmp4

		select Emp_ID,Emp_Name,join_date,
		Tot_Availed_Leave,celing,leave_type,
		Join_Yr_Leave,balance=(Join_Yr_Leave-Tot_Availed_Leave)  
		into #tmp4 from #tmp3

--		drop table #tmp5

		select * into #tmp5 from #tmp4

		union
		select Emp_id=(select Emp_id from #tmp4 group by Emp_id),
		Emp_name=(select Emp_name from #tmp4 group by Emp_name),
		Join_date=(select Join_date from #tmp4 group by Join_date),
		Tot_Availed_Leave=0,
		a.Celing,a.Leave_type,
		Join_Yr_Leave=0,Balance=a.Celing
		from lv_setup a
		where not exists (select * from #tmp4 b where a.Leave_type=b.Leave_type)
		------------------------------------------------
		select Emp_ID,Emp_Name,join_date,Tot_Availed_Leave,celing,leave_type,
		Join_Yr_Leave,balance =
		(balance-isnull((select leave from leave_as_cash where 
		leave_type= #tmp5.leave_type),0))
		from #tmp5		
				
end
--end ok 3-----------------------------------------------------------------

-----OK 4------------------------------

if @Yr_Duration > 1
begin
	declare @Yr_Duration1 as money
	set @Yr_Duration1=@Yr_Duration-1
		
	declare @EdDate1 as datetime
	set @EdDate1=(select getdate())

--	drop table #DT1 --Emp Total Leave without sum

	select a.Emp_ID,e.Emp_Name,e.join_date,a.Leave_st_Date,a.Leave_ed_date,
	a.Leave_Type,a.Total_Leave
	into #DT1 from leave a,emp_info e
	where e.Emp_ID=a.Emp_ID and a.Emp_ID=@emp_id
	and e.join_date between e.join_date and @EdDate1
	-----------------------------------------------
--	drop table #DT2 

	select t.Emp_ID,t.Emp_Name,t.join_date,t.Leave_Type,
	sum_leave=sum(t.Total_Leave),s.celing into #DT2 from #DT1 t,lv_setup s
	where t.leave_type=s.leave_type
	group by t.Emp_ID,t.Emp_Name,t.join_date,t.Leave_Type,s.celing

--	drop table #DT3

	select t.Emp_ID,t.Emp_Name,t.join_date,
	t.sum_leave Tot_Availed_Leave,t.celing,j.leave_type,j.Join_Yr_Leave 
	into #DT3 from #DT2 t,#JoinLeave j where j.leave_type=t.Leave_Type
	
--	drop table #DT4

	select Emp_ID,Emp_Name,join_date,
	Tot_Availed_Leave,celing,leave_type,
	Join_Yr_Leave,balance=(Join_Yr_Leave+(celing*@Yr_Duration1)-Tot_Availed_Leave)	
	into #DT4 from #DT3

--	Drop table #DT5

	select * into #DT5 from #DT4
	union
	select Emp_id=(select Emp_id from #dt4 group by Emp_id),
	Emp_name=(select Emp_name from #dt4 group by Emp_name),
	Join_date=(select Join_date from #dt4 group by Join_date),
	Tot_Availed_Leave=0,
	a.Celing,a.Leave_type,
	Join_Yr_Leave=0,Balance=a.Celing
	from lv_setup a
	where not exists (select * from #DT4 b where a.Leave_type=b.Leave_type)

--	Drop table #DT6

	select
	Emp_ID,Emp_Name,join_date,Tot_Availed_Leave,celing,leave_type,
	Join_Yr_Leave,balance =
	(balance-isnull((select leave from leave_as_cash where 
	leave_type= #dt5.leave_type),0))
	into #dt6
	from #dt5
	select*from #dt6
end
----end ok 4----------------------------

--end

set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SELECT_Leave    Script Date: 29/04/2002 12:40:15 AM ******/

/****** Object:  Stored Procedure dbo.SELECT_Leave    Script Date: 24/04/2002 05:16:19 PM ******/

/****** Object:  Stored Procedure dbo.SELECT_Leave    Script Date: 20/04/2002 08:35:30 PM ******/
CREATE PROCEDURE SELECT_Leave

@status varchar(1),
@Emp_ID varchar (10),
@Leave_st_Date datetime,
@Leave_ed_date datetime,
@Leave_Type varchar (15)

AS


if @status='1'
begin
		select * from leave where Emp_ID=@Emp_ID and Leave_st_Date=@Leave_st_Date 
		and Leave_ed_date=@Leave_ed_date and Leave_Type=@Leave_Type


end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  PROCEDURE S_Name_Select1

@status int,
@pat_id int,
@m_code varchar(2),
@s_code varchar(3)

AS
if @status=1
begin

select t.s_name from  pat_info_sub1 b,test_info_sub t
where b.m_code=t.m_code and b.s_code=t.s_code and b.pat_id=@pat_id
and b.m_code=@m_code and b.s_code=@s_code

end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO









/****** Object:  Stored Procedure dbo.S_name_select    Script Date: 26/09/2001 9:06:52 AM ******/
CREATE  PROCEDURE S_name_select

@Status int,
@m_code varchar(2),
@s_code varchar(3)

 AS

if @Status=1

begin
         select top 1 s_name from test_info_sub where m_code=@m_code and s_code=@s_code
end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE     PROCEDURE Search_Leave_Type

@status int

 AS

if @status=1

begin

	select leave_type from lv_setup

end

if @status=2

begin

select emp_name from emp_info order by emp_id

end

if @status=3

begin

	select a.item_name from item_info a,stock_out b where a.item_code=b.item_code order by a.item_code

end

if @status=4

begin

	select item_name from item_info order by item_name

end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




/****** Object:  Stored Procedure dbo.Search_Leave_as_cash    Script Date: 29/04/2002 12:40:15 AM ******/

/****** Object:  Stored Procedure dbo.Search_Leave_as_cash    Script Date: 24/04/2002 05:16:15 PM ******/
CREATE PROCEDURE Search_Leave_as_cash

@status int,
@emp_id varchar(10),
@leave_type varchar(15),
@cash_date datetime

AS

if @status=1
begin 
	select * from Leave_as_Cash where emp_id=@emp_id and leave_type=@leave_type and cash_date=@cash_date
end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE    PROC Search_Pat_ID

@mode int,
@pat_id1 varchar(50),
@refer_type varchar(3)
As

if @mode=1
begin
--select pat_id from pat_info_main where pat_id1=@pat_id1 and refer_type=@refer_type
declare @pat_id2 int
set @pat_id2=(select pat_id=isnull(sum(pat_id),0) 
from pat_info_main where pat_id1=@pat_id1 and refer_type=@refer_type)
select @pat_id2 pat_id2

end


if @mode=2
begin

select * from pat_info_main where pat_id1=@pat_id1

end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE    PROC Search_Pat_ID1

@mode int,
@pat_id1 varchar(50)
As

if @mode=1
begin

declare @pat_id2 int
set @pat_id2=(select pat_id=isnull(sum(pat_id),0) 
from pat_info_main where pat_id1=@pat_id1)
select @pat_id2 pat_id2

end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE     PROC Search_Pat_ID2

@mode int,
@pat_id1 varchar(50),
@refer_type varchar(3)
As

if @mode=1

begin



declare @pat_id int
set @pat_id=(select pat_id from pat_info_main where pat_id1=@pat_id1)
--select @pat_id pat_id

declare @tot_test_rate money
set @tot_test_rate=(select tot_test_rate=sum(test_rate) from pat_info_sub1 where pat_id=@pat_id)
--select @tot_test_rate tot_test_rate

declare @tot_VAT money
set @tot_VAT=(select vat_amt from pat_info_main where pat_id=@pat_id)
--select @tot_VAT tot_VAT
--select vat_amt from pat_info_main
declare @tot_collect_fee money
set @tot_collect_fee=(select sum(collect_fee) from pat_info_sub2 where pat_id=@pat_id)
--select @tot_collect_fee tot_collect_fee

declare @tot_Adv money
set @tot_Adv=(select sum(adv) from pat_info_sub2 where pat_id=@pat_id)
--select @tot_Adv tot_Adv


declare @tot_disc money
set @tot_disc=(select sum(disc) from pat_info_sub3 where pat_id=@pat_id)
--select @tot_disc tot_disc

declare @Tot_amount money

set @Tot_amount=(@tot_test_rate+@tot_VAT+@tot_collect_fee)
--select @Tot_amount Tot_amount

declare @status int
set @status=0

declare @due money

set @due=(@tot_test_rate+@tot_VAT+@tot_collect_fee-@tot_Adv-@tot_disc)
--select @due due
--for paid
if @due<=0
begin
	set @status=2
end

--for due
if @due>0
begin
	set @status=3
end

--for complimentary
if @Tot_amount=@tot_disc
begin
	set @status=1
end


select @status status



end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROC Search_Pat_Type

@mode int,
@pat_id1 varchar(50)
As
if @mode=1
begin

	select Row_Count=count(pat_id1),pat_id1 
	from pat_info_main where pat_id1=@pat_id1
	group by pat_id1

end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.Select_New_Doc_Name    Script Date: 08/11/2001 6:17:11 PM ******/
CREATE PROCEDURE Select_New_Doc_Name

@status int,
@StDate datetime,
@EdDate datetime

 AS

If @status=1

begin
	select  distinct doc_name from doctor_info_new where dt between @StDate and @EdDate

end



















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




/****** Object:  Stored Procedure dbo.Select_Paid    Script Date: 29/04/2002 12:40:15 AM ******/

/****** Object:  Stored Procedure dbo.Select_Paid    Script Date: 24/04/2002 05:16:20 PM ******/

/****** Object:  Stored Procedure dbo.Select_Paid    Script Date: 20/04/2002 08:35:30 PM ******/

/****** Object:  Stored Procedure dbo.Select_Paid    Script Date: 15/04/2002 11:52:59 AM ******/

/****** Object:  Stored Procedure dbo.Select_Paid    Script Date: 03/04/2002 4:10:16 PM ******/
CREATE PROCEDURE Select_Paid

@Status int,
@pat_id int

 AS

if @Status=1

begin
	select pat_id from pat_info_sub3 where pat_id=@pat_id and paid=1
end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO





/****** Object:  Stored Procedure dbo.Select_Soft_Sucurity    Script Date: 29/04/2002 12:40:15 AM ******/

/****** Object:  Stored Procedure dbo.Select_Soft_Sucurity    Script Date: 24/04/2002 05:16:20 PM ******/

/****** Object:  Stored Procedure dbo.Select_Soft_Sucurity    Script Date: 20/04/2002 08:35:31 PM ******/

/****** Object:  Stored Procedure dbo.Select_Soft_Sucurity    Script Date: 15/04/2002 11:52:55 AM ******/
CREATE  PROCEDURE [Select_Soft_Sucurity] 

@status int,
@uid varchar (10),
@screen_name varchar (50)

AS
if @status=1
begin
	select*from soft_security
	where uid=@uid and screen_name=@screen_name
end

if @status=2
begin
	select u_id from micropass
end

if @status=3
begin
	select scr_no from soft_bag
end
--select scr_no from soft_bag



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO




/****** Object:  Stored Procedure dbo.Select_Soft_Sucurity1    Script Date: 29/04/2002 12:40:15 AM ******/

/****** Object:  Stored Procedure dbo.Select_Soft_Sucurity1    Script Date: 24/04/2002 05:16:20 PM ******/

/****** Object:  Stored Procedure dbo.Select_Soft_Sucurity1    Script Date: 20/04/2002 08:35:31 PM ******/

/****** Object:  Stored Procedure dbo.Select_Soft_Sucurity1    Script Date: 15/04/2002 11:52:59 AM ******/
CREATE PROCEDURE [Select_Soft_Sucurity1] 

@status int,
@uid varchar (10),
@screen_name varchar (50)

AS
if @status=1
begin
	select allow from soft_security where uid=@uid and screen_name=@screen_name
end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE    PROCEDURE Select_VAT_Pat 

@status int,
@StDate datetime,
@EdDate datetime

AS
set nocount on

if @status=1
begin
	
---part 1--
--drop table #pat_info_sub1
select pat_id,test_rate=sum(test_rate) into #pat_info_sub1 
from pat_info_sub1 group by pat_id

--end part-1
---part 2--
--drop table #Final
	select a.pat_id,a.pat_id1,a.pat_name Patient_Name,a.refer_code,
	a.vat_amt,s.test_rate,Amount=(a.vat_amt+s.test_rate),
	b.doc_name Doctor_Name 
	into #Final from pat_info_main a,doctor_info b,#pat_info_sub1 s
	where a.refer_code=b.refer_code and a.pat_id=s.pat_id
	and refer_type='0' 
	and dt1 between @StDate and @EdDate
	select pat_id,pat_id1,Patient_Name,Doctor_Name,Amount from #Final order by pat_id
--end part 2--
end

if @status=2
begin
	
---part 1--
--drop table #pat_info_sub1N
	select pat_id,test_rate=sum(test_rate) into #pat_info_sub1N 
	from pat_info_sub1_VAT group by pat_id

--end part-1
---part 2--
--drop table #Final1
	select a.pat_id,a.pat_name Patient_Name,a.refer_code,
	a.vat_amt,s.test_rate,Amount=(a.vat_amt+s.test_rate),
	b.doc_name Doctor_Name 
	into #Final1 from pat_info_main_VAT a,doctor_info b,#pat_info_sub1N s
	where a.refer_code=b.refer_code and a.pat_id=s.pat_id
	and refer_type='0' 
	and dt1 between @StDate and @EdDate
	select pat_id,Patient_Name,Doctor_Name,Amount from #Final1
--end part 2--

end

set nocount off
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE St_Balance

@status int
As

SET NOCOUNT ON

if @status=1

begin

--drop table #Stk_Balance

	CREATE TABLE [#Stk_Balance] (
		[Item_code] [varchar] (45) NULL ,
		[Name_of_reagent] [varchar] (45) NULL ,
		[pur_qty] [int] NULL ,
		[itm_issue] [int] NULL
	)
	insert into #Stk_Balance
	select distinct p.item_code,i.item_name,qty=sum(p.test_per_box),0 from stock_in p,item_info i
	where i.item_code=p.item_code
	group by p.item_code,i.item_name
	
	
	
	
	insert into #Stk_Balance
	select distinct a.item_code,b.item_name,0,qty=sum(a.item_qty) from stock_out a,item_info b
	where a.item_code=b.item_code
	group by a.item_code,b.item_name
	
	select Item_code,Name_of_reagent,pur_qty=sum(pur_qty),itm_issue=sum(itm_issue),
	Balance=(sum(pur_qty)-sum(itm_issue)) from #Stk_Balance
	group by Item_code,Name_of_reagent

END

SET NOCOUNT OFF


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROCEDURE Stock_In_IUD
@Status varchar (1),
@inv_no varchar(25),
@item_code varchar (10),
@sup_id varchar (10),
@no_of_box int ,
@test_per_box int,
@amount money,
@pur_date datetime,
@exp_dt datetime,
@u_id varchar (10)

As

if @Status='I'
begin
	insert into Stock_In(inv_no,item_code,sup_id,no_of_box,test_per_box,amount,pur_date,exp_dt,u_id)
	values(@inv_no,@item_code,@sup_id,@no_of_box,@test_per_box,@amount,@pur_date,@exp_dt,@u_id)
end

if @Status='U'
begin

	update Stock_In set inv_no=@inv_no,item_code=@item_code,sup_id=@sup_id,no_of_box=@no_of_box,test_per_box=@test_per_box,
	amount=@amount,pur_date=@pur_date,exp_dt=@exp_dt,u_id=u_id where  inv_no=@inv_no and item_code=@item_code

end

if @Status='D'
begin

	delete from Stock_In where inv_no=@inv_no and item_code=@item_code

end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Stock_Out_IUD    Script Date: 29/04/2002 12:40:15 AM ******/

/****** Object:  Stored Procedure dbo.Stock_Out_IUD    Script Date: 24/04/2002 05:16:20 PM ******/

/****** Object:  Stored Procedure dbo.Stock_Out_IUD    Script Date: 20/04/2002 08:35:31 PM ******/
CREATE PROCEDURE Stock_Out_IUD
@status varchar(2),
@out_no varchar(25),
@item_code varchar (10),
@item_qty int ,
@emp_id varchar (10),
@notes varchar (100),
@issu_date datetime ,
@u_id varchar (10)

As

if @status='I'
begin
	insert into Stock_Out(out_no,item_code,item_qty,emp_id,notes,issu_date,u_id)
	values(@out_no,@item_code,@item_qty,@emp_id,@notes,@issu_date,@u_id)
end
if @status='U'
begin
	update Stock_Out set out_no=@out_no,item_code=@item_code,item_qty=@item_qty,emp_id=@emp_id,
	notes=@notes,issu_date=@issu_date,u_id=@u_id where out_no=@out_no and item_code=@item_code

end

if @status='D'
begin
	delete from Stock_Out where out_no=@out_no and item_code=@item_code

end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Stock_Status    Script Date: 29/04/2002 12:40:16 AM ******/

/****** Object:  Stored Procedure dbo.Stock_Status    Script Date: 24/04/2002 05:16:20 PM ******/

/****** Object:  Stored Procedure dbo.Stock_Status    Script Date: 20/04/2002 08:35:31 PM ******/
CREATE PROCEDURE Stock_Status

@Status int,
@item_code varchar(10),
@StDate datetime,
@Eddate datetime

AS
set nocount on
/*
if @Status=1
begin

	select u.item_code,a.item_name,Sum_item_qty=sum(u.item_qty),u.emp_id,e.emp_name,
	u.issu_date,Sum_test_per_box=sum(i.test_per_box) 
	from stock_out u,stock_in i,item_info a,emp_info e
	where i.item_code=u.item_code and a.item_code=u.item_code and e.emp_id=u.emp_id
	and u.issu_date between @StDate and @EdDate
	group by u.item_code,a.item_name,u.emp_id,e.emp_name,u.issu_date

end
*/
if @Status=1
begin
/*
	select u.item_code,a.item_name,Sum_item_qty=sum(u.item_qty),u.emp_id,e.emp_name,
	u.issu_date,Sum_test_per_box=sum(i.test_per_box) 
	from stock_out u,stock_in i,item_info a,emp_info e
	where u.item_code=@item_code 
	and i.item_code=u.item_code and a.item_code=u.item_code and e.emp_id=u.emp_id
	and u.issu_date between @StDate and @EdDate
	group by u.item_code,a.item_name,u.emp_id,e.emp_name,u.issu_date
*/

--	drop table #stock_in
	select item_code,sum_box=sum(no_of_box),sum_per_box=sum(test_per_box)
	into #stock_in from stock_in 
	where item_code=@Item_code and pur_date < @StDate group by item_code
--drop table #stock_in1
	select item_code,Pre_sum_box=sum_box,Pre_sum_per_box=sum_per_box,Pre_Qty=(sum_box*sum_per_box) into #stock_in1 from #stock_in
--	select Pre_sum_box=isnull(sum_box,0),Pre_sum_per_box=isnull(sum_per_box,0),
--	Pre_Qty=isnull((sum_box*sum_per_box),0) into #stock_in1 from #stock_in
--select*from #stock_in1

-----END patr 1--------

---part 2------------
--drop table #tmp1
SELECT i.inv_no,i.sup_id,s.sup_name,i.item_code,f.item_name,i.no_of_box,i.test_per_box,
Qty=(i.no_of_box*i.test_per_box),i.pur_date,i.exp_dt 
into #tmp1 from stock_in i,sup_info s,item_info f
where i.item_code=@Item_code and s.sup_id=i.sup_id and i.item_code=f.item_code
and i.pur_date between @StDate and @EdDate
-----------------end part 2-----
select a.inv_no,a.sup_id,a.sup_name,a.item_code,a.item_name,
a.no_of_box,a.test_per_box,a.Qty,a.pur_date,a.exp_dt,
b.Pre_sum_box,b.Pre_sum_per_box,b.Pre_Qty from #tmp1 a,#stock_in1 b









end

set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.Sup_Info_IUD    Script Date: 29/04/2002 12:40:16 AM ******/

/****** Object:  Stored Procedure dbo.Sup_Info_IUD    Script Date: 24/04/2002 05:16:20 PM ******/

/****** Object:  Stored Procedure dbo.Sup_Info_IUD    Script Date: 20/04/2002 08:35:31 PM ******/
CREATE   PROCEDURE Sup_Info_IUD


@Status varchar (1),
@sup_id varchar (10),
@sup_name varchar (50),
@sup_add varchar(100),
@sup_phone varchar(30),
@uid varchar(10)

AS


if @Status='I'
begin
	insert into sup_info(sup_id,sup_name,sup_add,sup_phone,uid)
	values(@sup_id,@sup_name,@sup_add,@sup_phone,@uid)
end

if @Status='U'
begin
	update sup_info set sup_id=@sup_id,sup_name=@sup_name,
	sup_add=@sup_add,sup_phone=@sup_phone,uid=@uid
	where sup_id=@sup_id
end

if @Status='D'
begin
	delete from sup_info where sup_id=@sup_id
end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROC Test_Result_Select10

@pat_id varchar(20),
@test_name varchar(500)
AS

--declare @pat_id varchar(20)
--declare @test_name varchar(500)
--set @pat_id=9462
--set @test_name='TOTAL COUNT'
select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_name=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code
and t.ref_range<>''




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROC Test_Result_Select11

@pat_id varchar(20),
@test_name varchar(500)
AS

--declare @pat_id varchar(20)
--declare @test_name varchar(500)
--set @pat_id=9462
--set @test_name='TOTAL COUNT'
select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_name=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code
and t.others<>''




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE   PROC Test_Result_Select12

@pat_id varchar(20),
@test_name varchar(500)
AS

--declare @pat_id varchar(20)
--declare @test_name varchar(500)
--set @pat_id=9462
--set @test_name='TOTAL COUNT'
select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_name=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code
--and t.ref_range<>''




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE    PROC Test_Result_Select13

@pat_id varchar(20),
@test_name varchar(500)
AS

--declare @pat_id varchar(20)
--declare @test_name varchar(500)
--set @pat_id=9462
--set @test_name='TOTAL COUNT'
select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_name=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code
and t.others1<>''



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE   PROC Test_Result_Select15

@pat_id varchar(20),
@test_name varchar(500)
AS

--declare @pat_id varchar(20)
--declare @test_name varchar(500)
--set @pat_id=9462
--set @test_name='TOTAL COUNT'
select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_name=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code
and t.others1<>''


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE    PROC Test_Result_Select16

@pat_id varchar(20),
@test_name varchar(500)
AS

--declare @pat_id varchar(20)
--declare @test_name varchar(500)
--set @pat_id=9462
--set @test_name='TOTAL COUNT'

select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_name=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code
and t.ref_range<>''



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create    PROC Test_Result_Select17

@pat_id varchar(20),
@test_name varchar(500)
AS


select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_name=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code
and t.ref_range<>''



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE PROC Test_Result_Select18

@pat_id varchar(20),
@test_name varchar(500)
AS

--declare @pat_id varchar(20)
--declare @test_name varchar(500)
--set @pat_id=9462
--set @test_name='TOTAL COUNT'
select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_name=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code
and t.others<>''




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE   PROC Test_Result_Select19

@pat_id varchar(20),
@test_name varchar(500)
AS

--declare @pat_id varchar(20)
--declare @test_name varchar(500)
--set @pat_id=9462
--set @test_name='TOTAL COUNT'
select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_result=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code
and t.ref_range<>''





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROC Test_Result_Select8

@pat_id varchar(20),
@test_name varchar(500)
AS

--declare @pat_id varchar(20)
--declare @test_name varchar(500)
--set @pat_id=9462
--set @test_name='TOTAL COUNT'
select t.test_name,t.test_result,t.unit,t.ref_range,others=isnull(t.others,''),others1=isnull(t.others1,'')
from test_result t,pat_info_sub1 a 
where a.pat_id=@pat_id and t.test_name=@test_name and 
a.s_code=t.s_code and a.m_code=t.m_code




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  PROCEDURE U_PAT_Disc

@status char(2),
@Disc money,
@unique_id int

AS


if @status='U'
begin
	
	update pat_info_sub3 set disc=@Disc
	where track_id=@unique_id

end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE U_PAT_Pay

@status char(2),
@adv money,
@collect_fee money,
@unique_id int

AS


if @status='U'
begin
	
	update pat_info_sub2 set adv=@adv,collect_fee=@collect_fee
	where unique_id=@unique_id

end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE  PROCEDURE U_PAT_Test_Code

@status char(2),
@m_code varchar(5),
@s_code varchar(5),
@type varchar(5),
@unique_id int

AS


if @status='UT'
begin
	
update pat_info_sub1 set m_code=@m_code,s_code=@s_code,type=@type
where unique_id=@unique_id

end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE  PROCEDURE U_VAT_ID
As

declare @Count as int
declare @var1 as int
declare @c as char(1)

DECLARE abc CURSOR FOR
SELECT distinct pat_id FROM pat_info_main_vat
OPEN abc

set @Count=1
set @c=1
WHILE @c = 1
begin
	FETCH NEXT FROM abc into @var1
	if @@FETCH_STATUS = 0 
	begin	
		update pat_info_main_vat set pat_id=@Count where pat_id=@var1
		set @count=@count+1

	end

	else set @c = 0
end
CLOSE abc
DEALLOCATE abc




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE Vat_Setup
As
INSERT INTO Vat_Per_Setup(vat_per)
VALUES(2.25)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




CREATE  PROCEDURE Vat_Setup_U
@mode int,
@vat_per money
As
set nocount on
if @mode=1
begin
	UPDATE Vat_Per_Setup
	SET vat_per=@vat_per
end

select message='Operated Successfully'

set nocount off

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.commission_pay_flush    Script Date: 15/10/2001 6:26:42 PM ******/
CREATE PROCEDURE commission_pay_flush

@Status int,
@refer_code varchar(10),
@s_code varchar(3)

 AS
if @status=1

select refer_code,pat_id,s_code,paid,dt,cleared,note 
from commission_pay where refer_code=@refer_code and s_code=@s_code








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.commission_select2    Script Date: 15/10/2001 6:26:40 PM ******/
CREATE PROCEDURE commission_select2

@status int,
@refer_code varchar(10) 

AS
if @status=1
begin
select*from commission_details where refer_code=@refer_code
end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.doc_comm_pay    Script Date: 12/11/2001 1:39:00 PM ******/
CREATE PROCEDURE doc_comm_pay

@Status int,
@refer_code varchar(10),
@pat_id int


AS
if @status=1

begin


select distinct a.pat_id,a.dt,Refer_code=@refer_code,
b.m_code,b.s_code,b.test_rate,a.vat_amt,c.disc,c.paid
from pat_info_main a,pat_info_sub1 b,pat_info_sub3 c,commission_per d,
test_info_main e,test_info_sub f
where a.pat_id=b.pat_id and
a.pat_id=c.pat_id and
d.refer_code=@refer_code and
b.m_code=e.m_code and
b.s_code=f.s_code and
c.paid=1 and a.pat_id not in (select pat_id from commission_pay)

end

if @status=2
begin
	select refer_code,pat_id,m_code,s_code,paid,test_rate,vat_amt,disc,dt from commission_pay where refer_code=@refer_code
	and pat_id=@pat_id
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


/****** Object:  Stored Procedure dbo.doc_comm_pay1    Script Date: 15/11/2001 12:26:19 PM ******/
CREATE PROCEDURE doc_comm_pay1

@Status char(2) ,
@refer_code varchar(10),
@pat_id int

AS
if @status='I'

begin

insert into commission_pay (pat_id,dt,refer_code,m_code,s_code,test_rate,vat_amt,disc,paid)
select distinct a.pat_id,a.dt,Refer_code=@refer_code,
b.m_code,b.s_code,b.test_rate,a.vat_amt,c.disc,c.paid
from pat_info_main a,pat_info_sub1 b,pat_info_sub3 c,commission_per d,
test_info_main e,test_info_sub f
where a.pat_id=b.pat_id and
a.pat_id=c.pat_id and
d.refer_code=@refer_code and
b.m_code=e.m_code and
b.s_code=f.s_code and
c.paid=1 and 
a.pat_id=@pat_id
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


/****** Object:  Stored Procedure dbo.doc_comm_pay2    Script Date: 15/11/2001 12:26:19 PM ******/
CREATE PROCEDURE doc_comm_pay2

@Status int,
@refer_code varchar(10),
@pat_id int

AS
if @status=1

begin

select distinct a.pat_id,a.pat_name,a.dt,Refer_code=@refer_code,
Doc_Name=(select doc_name from doctor_info where refer_code=@refer_code),
b.m_code,b.s_code,b.test_rate,a.vat_amt,c.disc,d.comm_per,
e.m_name,f.s_name
from pat_info_main a,pat_info_sub1 b,pat_info_sub3 c,commission_per d,
test_info_main e,test_info_sub f
where a.pat_id=b.pat_id and
a.pat_id=c.pat_id and
d.refer_code=@refer_code and
b.m_code=e.m_code and
b.s_code=f.s_code and
c.paid=1 and a.pat_id=@pat_id and a.pat_id not in (select pat_id from commission_pay where pat_id=@pat_id)
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



/****** Object:  Stored Procedure dbo.doc_comm_payU    Script Date: 15/11/2001 12:26:19 PM ******/
CREATE PROCEDURE doc_comm_payU

@Status char(2) ,
@refer_code varchar(10),
@pat_id int

AS
if @status='U'

begin

insert into commission_pay (pat_id,dt,refer_code,m_code,s_code,test_rate,vat_amt,disc,paid)
select distinct a.pat_id,a.dt,Refer_code=@refer_code,
b.m_code,b.s_code,b.test_rate,a.vat_amt,c.disc,c.paid
from pat_info_main a,pat_info_sub1 b,pat_info_sub3 c,commission_per d,
test_info_main e,test_info_sub f
where a.pat_id=b.pat_id and
a.pat_id=c.pat_id and
d.refer_code=@refer_code and
b.m_code=e.m_code and
b.s_code=f.s_code and
c.paid=1 and 
a.pat_id=@pat_id
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.m_name_select    Script Date: 08/11/2001 6:17:02 PM ******/
/****** Object:  Stored Procedure dbo.m_name_select    Script Date: 26/09/2001 9:06:45 AM ******/
CREATE PROCEDURE m_name_select

@Status int,
@m_name varchar(40)

 AS

if @Status=1

begin
	select*from test_info_main where m_name=@m_name
end

if @Status=2

begin
	select*from test_info_main where m_name=@m_name
end

if @Status=3

begin
	select*from doctor_info where refer_code=@m_name
end















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pass_para    Script Date: 29/04/2002 12:40:16 AM ******/

/****** Object:  Stored Procedure dbo.pass_para    Script Date: 24/04/2002 05:16:21 PM ******/
CREATE PROCEDURE pass_para

@Mode char (1),
@Id varchar (20),
@M_code varchar (2),
@S_code varchar (3),
@Doc_Path varchar (200) 

AS

if @Mode = 'I'
begin
	insert into vbw values(@Id,@M_code,@S_code,@Doc_Path+@M_code+@S_code+@Id+'.rtf')
end

if @Mode = 'B'
begin
	select * from vbw where id=@Id and M_code=@M_code and S_code=@S_code
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







/****** Object:  Stored Procedure dbo.pro_COMPANY_INFO    Script Date: 26/09/2001 9:06:47 AM ******/
CREATE   PROCEDURE pro_COMPANY_INFO

@status char(1),

@comp_name varchar (50),
@addr varchar (200),
@uid varchar (20),
@dt datetime 

AS

if @status='I'
begin
	insert into company_info(comp_name,addr,uid,dt)
	values(@comp_name,@addr,@uid,@dt) 

end

if @status='U'
begin
	update company_info set comp_name=@comp_name,addr=@addr,uid=@uid,dt=@dt where comp_name=@comp_name

end

if @status='D'
begin
	delete from company_info where comp_name=@comp_name	

end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.pro_COMMISSION_MAIN    Script Date: 26/09/2001 9:06:46 AM ******/
CREATE PROCEDURE pro_Commission_Details

@status char(1),
@refer_code varchar(10),
@pat_id varchar (10),
@m_code varchar (2),
@s_code varchar (3),
@commission money,
@uid varchar (10),
@dt datetime,
@note varchar(500)

AS

if @status='I'
begin
	insert into Commission_Details(refer_code,pat_id,m_code,s_code,commission,uid,dt,note)
	values(@refer_code,@pat_id,@m_code,@s_code,@commission,@uid,@dt,@note)
end

if @status='U'
begin	
	update Commission_Details set refer_code=@refer_code,pat_id=@pat_id,m_code=@m_code,s_code=@s_code,commission=@commission,uid=@uid,dt=@dt,note=@note 
	where  refer_code=@refer_code and pat_id=@pat_id and m_code=@m_code and s_code=@s_code
end

if @status='D'
begin
	delete from Commission_Details where refer_code=@refer_code and pat_id=@pat_id and m_code=@m_code and s_code=@s_code
end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_COMMISSION_SUB    Script Date: 26/09/2001 6:32:14 PM ******/


CREATE PROCEDURE pro_Commission_Pay

@status char(1),
@refer_code varchar(10),
@pat_id varchar(10),
@m_code varchar(2),
@s_code varchar(3),
@paid money,
@uid varchar (10),
@dt datetime ,
@cleared varchar(1),
@note varchar(500)

AS

if @status='I'
begin
	insert into Commission_Pay(refer_code,pat_id,m_code,s_code,paid,uid,dt,cleared,note)
	values(@refer_code,@pat_id,@m_code,@s_code,@paid,@uid,@dt,@cleared,@note)
end

if @status='U'
begin	

	update Commission_Pay set refer_code=@refer_code,
             pat_id=@pat_id,m_code=@m_code,s_code=@s_code,paid=@paid,uid=@uid,dt=@dt,cleared=@cleared,note=@note
end

if @status='D'
begin
	delete from Commission_Pay where refer_code=@refer_code and pat_id=@pat_id and m_code=@m_code and s_code=@s_code
--delete from commission_summary where refer_code=@refer_code,pat_id=@pat_id
end











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.pro_Commission_Per    Script Date: 20/01/2002 10:27:01 AM ******/

/****** Object:  Stored Procedure dbo.pro_Commission_Per    Script Date: 08/11/2001 6:17:04 PM ******/
CREATE PROCEDURE pro_Commission_Per

@status char(1),

@type varchar (10),
@refer_code varchar (10),
@comm_per money

AS

if @status='I'
begin
	insert into commission_per(type,refer_code,comm_per)
	values(@type,@refer_code,@comm_per)

end

if @status='U'
begin	

	update Commission_per set type=@type,refer_code=@refer_code,comm_per=@comm_per
	where type=@type and refer_code=@refer_code
end

if @status='D'
begin

	delete from Commission_per where type=@type and refer_code=@refer_code

end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.pro_DOCTOR_INFO    Script Date: 29/04/2002 12:40:17 AM ******/

/****** Object:  Stored Procedure dbo.pro_DOCTOR_INFO    Script Date: 08/11/2001 6:17:04 PM ******/


/****** Object:  Stored Procedure dbo.pro_DOCTOR_INFO    Script Date: 26/09/2001 9:06:47 AM ******/
CREATE   PROCEDURE pro_DOCTOR_INFO

@status char(2),
@refer_code varchar (10),
@doc_name varchar (200),
@addr varchar (200),
@phone varchar (25),
@fax varchar (25),
@email varchar (25),
@birth_date datetime,
@marriage_date datetime,
@uid varchar (10),	
@EmpId VARCHAR (2)



AS

if @status='I'
begin
	insert into doctor_info(refer_code,doc_name,addr,phone,fax,email,birth_date,marriage_date,uid, EmpId)
	values(@refer_code,@doc_name,@addr,@phone,@fax,@email,@birth_date,@marriage_date,@uid,@EmpId)
end

if @status='U'
begin	
	update doctor_info set refer_code=@refer_code,doc_name=@doc_name,addr=@addr,
             phone=@phone,fax=@fax,email=@email,birth_date=@birth_date,marriage_date=@marriage_date,uid=@uid,EmpId=@EmpId where refer_code=@refer_code
end

if @status='D'
begin
	delete from doctor_info where refer_code=@refer_code
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_DOCTOR_INFO_NEW    Script Date: 08/11/2001 6:17:04 PM ******/
CREATE  PROCEDURE pro_DOCTOR_INFO_NEW

@status char(2),
@pat_id int,
@doc_name varchar (45),
@addr varchar (200),
@phone varchar (25),
@fax varchar (25),
@email varchar (25),
@uid varchar (10),
@doc_date datetime
AS

if @status='I'
begin
	insert into doctor_info_new(pat_id,doc_name,addr,phone,fax,email,uid,doc_date)
	values(@pat_id,@doc_name,@addr,@phone,@fax,@email,@uid,@doc_date)
end

if @status='U'
begin	
	update doctor_info_new set pat_id=@pat_id,doc_name=@doc_name,addr=@addr,
	phone=@phone,fax=@fax,email=@email,
	uid=@uid,doc_date=@doc_date where pat_id=@pat_id
end

--if @status='U1'
--begin	
--	update doctor_info_new set pat_id=@pat_id where pat_id=@pat_id
--end


if @status='D'
begin
	delete from doctor_info_new where pat_id=@pat_id
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO





/****** Object:  Stored Procedure dbo.pro_DOCTOR_INFO_NEW    Script Date: 08/11/2001 6:17:04 PM ******/
CREATE   PROCEDURE pro_DOCTOR_INFO_NEW1

@status char(2),
@pat_id int

AS


if @status='U'
begin	
	update doctor_info_new set pat_id=@pat_id where pat_id=0
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE   PROCEDURE pro_DOCTOR_INFO_NEW2

@status char(2),
@pat_id int,
@uid varchar(15)

AS


if @status='U'
begin	
	update doctor_info_new set pat_id=@pat_id where pat_id=0 and uid=@uid
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE pro_DOCTOR_INFO_NEW3

@pat_id int,
@doc_name varchar (45),
@addr varchar (200),
@phone varchar (25),
@fax varchar (25),
@email varchar (25),
@uid varchar (10),
@doc_date datetime
AS

IF NOT EXISTS(select*from doctor_info_new where pat_id='0' and uid=@uid)

begin
	insert into doctor_info_new(pat_id,doc_name,addr,phone,fax,email,uid,doc_date)
	values(@pat_id,@doc_name,@addr,@phone,@fax,@email,@uid,@doc_date)
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROCEDURE pro_Emp_Info

@status char(1),
@Emp_ID varchar (10),
@Emp_Name varchar(45),
@join_date datetime,
@Emp_Desig varchar (25),
@Title varchar (25),
@Salary money,
@Sex varchar(6),
@Age VARCHAR(20),
@Emp_Pre_Add varchar(50),
@Emp_Per_Add varchar (50),
@Emp_Phone varchar (40),
@Emp_Email varchar (25)


AS

if @status='I'
begin



insert into Emp_Info(Emp_ID,Emp_Name,join_date,Emp_Desig,Title,Salary,Sex,Age,Emp_Pre_Add,Emp_Per_Add,Emp_Phone,Emp_Email)
values(@Emp_ID,@Emp_Name,@join_date,@Emp_Desig,@Title,@Salary,@Sex,@Age,@Emp_Pre_Add,@Emp_Per_Add,@Emp_Phone,@Emp_Email)

end


if @status='U'
begin
	update Emp_Info set Emp_ID=@Emp_ID,Emp_Name=@Emp_Name,join_date=@join_date,Emp_Desig=@Emp_Desig,
	Title=@Title,Salary=@Salary,Sex=@Sex,Age=@Age,
	Emp_Pre_Add=@Emp_Pre_Add,Emp_Per_Add=@Emp_Per_Add,Emp_Phone=@Emp_Phone,
	Emp_Email=@Emp_Email
	 where Emp_ID=@Emp_ID

end

if @status='D'
begin
	delete from Emp_Info where Emp_ID=@Emp_ID

end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_Leave_IUD    Script Date: 29/04/2002 12:40:17 AM ******/

/****** Object:  Stored Procedure dbo.pro_Leave_IUD    Script Date: 24/04/2002 05:16:21 PM ******/

/****** Object:  Stored Procedure dbo.pro_Leave_IUD    Script Date: 20/04/2002 08:35:26 PM ******/
CREATE PROCEDURE pro_Leave_IUD

@status char(1),

@Emp_ID varchar (10),
@Leave_st_Date datetime ,
@Leave_ed_date datetime ,
@Leave_Type varchar (15),
@Total_Leave int


AS

if @status='I'
begin



	insert into leave(Emp_ID,Leave_st_Date,Leave_ed_date,Leave_Type,Total_Leave)
	values(@Emp_ID,@Leave_st_Date,@Leave_ed_date,@Leave_Type,@Total_Leave)

end


if @status='U'
begin

	update leave set Emp_ID=@Emp_ID,Leave_st_Date=@Leave_st_Date,Leave_ed_date=@Leave_ed_date,Leave_Type=@Leave_Type,
	Total_Leave=@Total_Leave where Emp_ID=@Emp_ID and Leave_st_Date=@Leave_st_Date 
	and Leave_ed_date=@Leave_ed_date and Leave_Type=@Leave_Type
end

if @status='D'
begin
	delete from leave where Emp_ID=@Emp_ID and Leave_st_Date=@Leave_st_Date 
	and Leave_ed_date=@Leave_ed_date and Leave_Type=@Leave_Type

end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE    PROCEDURE pro_Leave_setup_IUD

@status char(1),
@Leave_Type varchar (15),
@celing int,
@lv_st_year datetime

AS

if @status='I'
begin



	insert into lv_setup(Leave_Type,Celing,lv_st_year)
	values(@Leave_Type,@Celing,@lv_st_year)

end


if @status='U'
begin

	update lv_setup set celing=@Celing where Leave_type=@Leave_Type
	update lv_setup set lv_st_year=@lv_st_year

end

if @status='D'
begin
	delete  lv_setup  where Leave_type=@Leave_Type

end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.pro_PAT_INFO_MAIN    Script Date: 26/09/2001 9:06:48 AM ******/
CREATE   PROCEDURE pro_PAT_INFO_MAIN

@status char(1),

@pat_name varchar (45),
@sex varchar (6),
@age varchar (17),
@refer_code varchar (10),
@addr varchar (200),
@phone varchar (25),
@fax varchar (25),
@email varchar (25),
@uid varchar (10),
@dt datetime,
@vat_per money,
@vat_amt money,
@booth varchar(2),
@tmp_dt datetime,
@dt1 datetime,
@refer_type varchar (3),
@pat_id1 varchar(50),
@pat_MY varchar(4)

AS

if @status='I'
begin
	insert into pat_info_main(pat_name,sex,age,refer_code,addr,phone,fax,email,uid,dt,vat_per,vat_amt,booth,tmp_dt,dt1,refer_type,pat_id1,pat_MY)	
	values(@pat_name,@sex,@age,@refer_code,@addr,@phone,@fax,@email,@uid,@dt,@vat_per,@vat_amt,@booth,@tmp_dt,@dt1,@refer_type,@pat_id1,@pat_MY)

end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.pro_PAT_INFO_MAIN_UD    Script Date: 12/11/2001 1:39:04 PM ******/
CREATE    PROCEDURE pro_PAT_INFO_MAIN_UD

@status char(1),
@pat_id int,
@pat_name varchar (45),
@sex varchar (6),
@age varchar (17),
@refer_code varchar (10),
@addr varchar (200),
@phone varchar (25),
@fax varchar (25),
@email varchar (25),
@uid varchar (10),
@dt datetime,
@vat_per money,
@vat_amt money,
@booth varchar(2),
@tmp_dt datetime,
@dt1 datetime,
@refer_type varchar(3),
@pat_id1 varchar(50),
@pat_my varchar(4)

AS


if @status='U'
begin
	update pat_info_main set pat_name=@pat_name,sex=@sex,age=@age,refer_code=@refer_code,addr=@addr,phone=@phone,
             fax=@fax,email=@email,uid=@uid,dt=@dt,vat_per=@vat_per,vat_amt=@vat_amt,booth=@booth,tmp_dt=@tmp_dt,dt1=@dt1,refer_type=@refer_type,pat_id1=@pat_id1,pat_my=@pat_my where pat_id=@pat_id

end

if @status='D'
begin
	delete from pat_info_main where pat_id=@pat_id

end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_PAT_INFO_SUB1    Script Date: 26/09/2001 9:06:48 AM ******/
CREATE PROCEDURE pro_PAT_INFO_SUB1

@status char(1),

@pat_id int,
@m_code varchar (2),
@s_code varchar (3),
@test_rate money,
@delv_dt datetime,
@type varchar (10),
@uid varchar (10),
@dt datetime, 
@tmp_dt datetime,
@dt1 datetime,
@unique_id int
AS

if @status='I'
begin
	insert into pat_info_sub1(pat_id,m_code,s_code,test_rate,delv_dt,type,uid,dt,tmp_dt,dt1)
	values(@pat_id,@m_code,@s_code,@test_rate,@delv_dt,@type,@uid,@dt,@tmp_dt,@dt1) 
end

if @status='U'
begin
	update pat_info_sub1 set pat_id=@pat_id,m_code=@m_code,s_code=@s_code,test_rate=@test_rate,delv_dt=@delv_dt,type=@type,uid=@uid,dt=@dt,tmp_dt=@tmp_dt,dt1=@dt1
	where pat_id=@pat_id
end

if @status='D'
begin
	delete from pat_info_sub1 where unique_id=@unique_id
end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.pro_PAT_INFO_SUB2    Script Date: 26/09/2001 9:06:48 AM ******/
CREATE  PROCEDURE pro_PAT_INFO_SUB2

@status char(1),
@pat_id int,
@adv money,
@uid varchar (10),
@dt datetime,
@collect_fee money,
@type varchar (5),
@tmp_dt datetime,
@dt1 datetime,
@dt2 datetime,
@unique_id int

AS

if @status='I'
begin
	insert into pat_info_sub2(pat_id,adv,uid,dt,collect_fee,type,tmp_dt,dt1,dt2) 
	values(@pat_id,@adv,@uid,@dt,@collect_fee,@type,@tmp_dt,@dt1,@dt2)
end

if @status='U'
begin
	
	update pat_info_sub2 set pat_id=@pat_id,adv=@adv,uid=@uid,dt=@dt,collect_fee=@collect_fee,type=@type,tmp_dt=@tmp_dt,dt1=@dt1 where pat_id=@pat_id

end

if @status='D'
begin
	delete from pat_info_sub2 where unique_id=@unique_id
end










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_PAT_INFO_SUB3    Script Date: 26/09/2001 9:06:48 AM ******/
CREATE PROCEDURE pro_PAT_INFO_SUB3

@status char(1),

@pat_id int,
@disc money,
@paid money,
@uid varchar(10),
@dt datetime,
@tmp_dt datetime,
@dt1 datetime

AS

if @status='I'
begin
	insert into pat_info_sub3(pat_id,disc,paid,uid,dt,tmp_dt,dt1) 
	values(@pat_id,@disc,@paid,@uid,@dt,@tmp_dt,@dt1)
end

if @status='U'
begin
	update pat_info_sub3 set pat_id=@pat_id,disc=@disc,paid=@paid,uid=@uid,dt=@dt,tmp_dt=@tmp_dt,dt1=@dt1 where pat_id=@pat_id
end

if @status='D'
begin
	delete from pat_info_sub3 where pat_id=@pat_id
end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_Report_All    Script Date: 26/09/2001 9:06:49 AM ******/
CREATE  PROCEDURE pro_Report_All

@status char(1),

@pat_id varchar(10),
@m_code varchar (3),
@s_code varchar (3),
@field1 varchar (1500),
@field2 varchar (1500),
@field3 varchar (1500),
@field4 varchar (1500),
@field5 varchar (1500),
@field6 varchar (1500),
@field7 varchar (1500),
@field8 varchar (1500),
@field9 varchar (1500),
@field10 varchar (1500),
@field11 varchar (1500),
@field12 varchar (1500),
@field13 varchar (1500),
@field14 varchar (1500),
@field15 varchar (1500),
@uid varchar (10),
@dt datetime,
@type varchar (2),
@pat_id1 varchar(50)

AS
if @status='I'
begin
         insert into report_all(pat_id,m_code,s_code,field1,field2,field3,field4,field5,field6,field7,field8,field9,field10,field11,field12,field13,
         field14,field15,uid,dt,type,pat_id1)
         values(@pat_id,@m_code,@s_code,@field1,@field2,@field3,@field4,@field5,@field6,@field7,@field8,@field9,@field10,
                    @field11,@field12,@field13,@field14,@field15,@uid,@dt,@type,@pat_id1)
	
end

if @status='U'
begin	
	update report_all set pat_id=@pat_id,m_code=@m_code,s_code=@s_code,field1=@field1,field2=@field2,
             field3=@field3,field4=@field4,field5=@field5,field6=@field6,field7=@field7,field8=@field8,field9=@field9,
             field10=@field10,field11=@field11,field12=@field12,field13=@field13,field14=@field14,field15=@field15,uid=@uid,dt=@dt,type=@type,pat_id1=@pat_id1
	where pat_id=@pat_id and m_code=@m_code and s_code=@s_code and type=@type


end

if @status='D'
begin
	delete from report_all where pat_id=@pat_id and m_code=@m_code and s_code=@s_code
end
































GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO




/****** Object:  Stored Procedure dbo.pro_Soft_Security    Script Date: 29/04/2002 12:40:18 AM ******/

/****** Object:  Stored Procedure dbo.pro_Soft_Security    Script Date: 24/04/2002 05:16:22 PM ******/

/****** Object:  Stored Procedure dbo.pro_Soft_Security    Script Date: 20/04/2002 08:35:27 PM ******/

/****** Object:  Stored Procedure dbo.pro_Soft_Security    Script Date: 15/04/2002 11:53:01 AM ******/
CREATE PROCEDURE [pro_Soft_Security] 

@status varchar (2),
@uid varchar (15),
@user_type varchar (50),
@screen_name varchar (50),
@Allow varchar (3) 

AS
if @status='I'
begin
	insert into soft_security(uid,user_type,screen_name,Allow)
	values(@uid,@user_type,@screen_name,@Allow)
end

if @status='U'
begin
	update soft_security set uid=@uid,user_type=@user_type,screen_name=@screen_name,Allow=@Allow
	where uid=@uid and screen_name=@screen_name
end

if @status='D'
begin
	delete from soft_security where uid=@uid and screen_name=@screen_name
end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROCEDURE [pro_Soft_Security1] 
AS

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmUser_Authority','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmChange_Password','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmCommission_Per','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmCompany_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmCreate_User','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmDoctor_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmEmp_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmFont','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmItem_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmLeave','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmLeave_Setup','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmPatient_Info_VAT','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmPatient_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmStock_IN','YES')
-----------------------------------------------------------
insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmStock_Out','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmSup_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmTest_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmTest_Result','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmVAT_Setup','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmPat_Info_Due','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmPay_Edit','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','frmDisc_Edit','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rAdv_Coll','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rBio_Chamical','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rBody_Fluid','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rCT_SCAN','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rDrug','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rEchocardiography','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rEndoscopy','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rHaematology','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rHepatitis','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rHistopathology','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rHormone','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rImmunology','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rMammography','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rMicrobiology','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rPaps','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rStool','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rTumour_Marker','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rUltrasonogram','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rUrine1','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rX_Ray','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rBooth_User_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rDaily_Statement','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rDoc_Due_Pat','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rDoc_New','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rDoc_Pay','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rLeave_Balance','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rPat_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rPat_Type','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','RptDoctor_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','RptTest_Info','YES')

insert into soft_security(uid,user_type,screen_name,Allow)
values('dsl','Admin','rStock_Status','YES')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_TEST_INFO_MAIN    Script Date: 26/09/2001 9:06:49 AM ******/
CREATE  PROCEDURE pro_TEST_INFO_MAIN

@status char(2),
@m_code varchar (2),
@m_name varchar (40),
@uid varchar (10)
--@dt datetime

AS

if @status='I'
begin
	insert into test_info_main(m_code,m_name,uid)
	values(@m_code,@m_name,@uid)
end

if @status='U'
begin
	update test_info_main set m_code=@m_code,m_name=@m_name,uid=@uid where m_code=@m_code
end

if @status='D'
begin
	delete from test_info_sub where m_code=@m_code and s_code=@m_name
	delete from test_info_rate where m_code=@m_code and s_code=@m_name
end

if @status='D1'
begin
	delete from test_info_main where m_code=@m_code
end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.pro_TEST_INFO_RATE    Script Date: 26/09/2001 9:06:49 AM ******/
CREATE PROCEDURE pro_TEST_INFO_RATE

@status char(1),
@m_code varchar (2),
@s_code varchar (3),
@rate money,
@uid varchar (10)


AS

if @status='I'
begin
	insert into test_info_rate(m_code,s_code,rate,uid)
	values(@m_code,@s_code,@rate,@uid)  	
end

if @status='U'
begin
	update test_info_rate set m_code=@m_code,s_code=@s_code,rate=@rate,uid=@uid where m_code=@m_code and s_code=@s_code
end

if @status='D'
begin	
	delete from test_info_rate where s_code=@s_code and m_code=@m_code
end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.pro_TEST_INFO_SUB    Script Date: 16/10/2001 12:00:17 AM ******/

CREATE PROCEDURE pro_TEST_INFO_SUB

@status char(1),

@s_code varchar (3),
@s_name varchar (60),
@type varchar (25),
@m_code varchar (2),
@uid varchar (10)
--@dt datetime

AS

if @status='I'
begin
	
	insert into test_info_sub(s_code,s_name,type,m_code,uid)
	values(@s_code,@s_name,@type,@m_code,@uid)
end

if @status='U'
begin
	update test_info_sub set s_code=@s_code,s_name=@s_name,type=@type,m_code=@m_code,uid=@uid
	where s_code=@s_code and m_code=@m_code
end

if @status='D'
begin	
	delete from test_info_sub where s_code=@s_code and m_code=@m_code
end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE    PROCEDURE pro_TEST_RESULT

@status char(2),
@test_name varchar (1500),
@test_result varchar (1500),
@unit varchar (1500),
@ref_range varchar (1500),
@type varchar (4),
@others varchar(2000),
@others1 varchar(2000),
@m_code varchar(4),
@s_code varchar(4)



AS

if @status='I'
begin
	insert into test_result(test_name,test_result,unit,ref_range,type,others,others1,m_code,s_code)
	values(@test_name,@test_result,@unit,@ref_range,@type,@others,@others1,@m_code,@s_code)
end

if @status='U'
begin	
	update test_result set test_name=@test_name,test_result=@test_result,unit=@unit,ref_range=@ref_range,type=@type,
	others=@others,others1=@others1,m_code=@m_code,s_code=@s_code
	where test_name=@test_name and type=@type
end

if @status='D'
begin
	delete from test_result where type=@type and test_name=@test_name
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE   PROCEDURE pro_TEST_RESULT1

@status int,
@test_name varchar (500),
@test_result varchar (1500),
@unit varchar (1500),
@ref_range varchar (1500),
@type varchar (4),
@others varchar(2000),
@others1 varchar(2000),
@m_code varchar(2),
@s_code varchar(2),
@unique_id int



AS



if @status=1
begin	
	update test_result set test_name=@test_name,test_result=@test_result,unit=@unit,ref_range=@ref_range,type=@type,
	others=@others,others1=@others1,m_code=@m_code,s_code=@s_code
	where unique_id=@unique_id
end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE    PROCEDURE pro_Test_Info_FLUSH

@Status char (1),
@m_code varchar (2)
--@m_name varchar (40),
--@s_code varchar (3),
--@s_name varchar (40),
--@type varchar (25),
--@rate money

AS
set nocount on
if @Status='1'
begin

CREATE TABLE #TMP
	(m_code VARCHAR(2),
	m_name varchar(200),
	s_code varchar(4),
	s_name varchar(200),
	rate money,
	type varchar(20),
	s_code1 int
	)
insert into #TMP
select distinct a.m_code,a.m_name,b.s_code,b.s_name,c.rate,b.type,b.s_code from
test_info_main a,test_info_sub b,test_info_rate c where
a.m_code=b.m_code and a.m_code=c.m_code and b.s_code=c.s_code and a.m_code=@m_code

select * from #TMP order by s_code1



end

if @Status='2'
begin

select distinct a.m_code,a.m_name,b.s_code,b.s_name,c.rate,b.type from
test_info_main a,test_info_sub b, test_info_rate c where
a.m_code=b.m_code and a.m_code=c.m_code and b.s_code=c.s_code
end

if @Status='3'
begin

select scr_no Screen_Name,descript Screen_Description from soft_bag order by scr_no

end

if @Status='4'
begin

select a.uid U_ID,b.u_name,a.Screen_Name,c.descript Screen_Description,
a.Allow,a.User_Type from soft_security a,micropass b,soft_bag c
where a.uid=b.u_id and c.scr_no=a.screen_name order by b.u_name


end


set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_USR_INFO    Script Date: 26/09/2001 9:06:50 AM ******/
CREATE PROCEDURE pro_USR_INFO

@status char(1),

@usr_id varchar (10),
@usr_name varchar (45),
@desig varchar (25),
@degree varchar (50),
@uid varchar (10),
@dt datetime

AS

if @status='I'
begin
	insert into usr_info(usr_id,usr_name,desig,degree,uid,dt)
	values(@usr_id,@usr_name,@desig,@degree,@uid,@dt)

end

if @status='U'
begin	
	update usr_info set usr_id=@usr_id,usr_name=@usr_name,desig=@desig,degree=@degree,uid=@uid,dt=@dt where usr_id=@usr_id

end

if @status='D'
begin
	delete from usr_info where usr_id=@usr_id	

end



























GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







/****** Object:  Stored Procedure dbo.pro_commpay__select1    Script Date: 15/10/2001 6:26:44 PM ******/
CREATE PROCEDURE pro_commpay__select1 -- this PROCEDURE is not using now
-- this PROCEDURE is not using now
@status int,
@refer_code varchar(10),
@pat_id varchar(10),
@s_code varchar(3)

AS

if @status=1

begin

        select*from commission_pay where refer_code=@refer_code and pat_id=@pat_id and s_code=@s_code

end






















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_micropass    Script Date: 29/04/2002 12:40:18 AM ******/


/****** Object:  Stored Procedure dbo.pro_micropass    Script Date: 08/11/2001 6:17:05 PM ******/
CREATE PROCEDURE [pro_micropass]

@psl numeric,
@u_id varchar (15),
@u_name varchar (200),
@user_pass varchar (45),
@uid char (15),
@udt datetime,
@cancel bit,
@status char(1)
 
 AS


if @status='I'
begin
	insert into micropass (u_id,u_name,uid,udt,cancel) values (@u_id,@u_name,@uid,getdate(),@cancel)
end

if @status='U'
begin
	--insert into a_micropass
	select *,auid=@uid,audt=getdate(),delflg=0 from micropass where u_id=@u_id

	update micropass set
	u_name=@u_name,uid=@uid,udt=getdate(),cancel=@cancel where u_id=@u_id
end

if @status='D'
begin
	--insert into a_micropass
	select *,auid=@uid,audt=getdate(),delflg=1 from micropass where u_id=@u_id

	delete from micropass where u_id=@u_id
end

if @status='E'
begin
	--insert into a_micropass
	select *,auid=@uid,audt=getdate(),delflg=0 from micropass where u_id=@u_id

	update micropass set
	user_pass=@user_pass,uid=@uid,udt=getdate() where u_id=@u_id
end






















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE          PROCEDURE pro_name_SELECT
@status char(3),
@select_name varchar(50)

AS


if @status='1'
begin
		select doc_name from doctor_info where refer_code like @select_name 
--order by doc_name
end

if @status=2
begin

	select a.m_code,a.m_name,b.s_code,b.s_name,c.rate from
	test_info_main a,test_info_sub b, test_info_rate c where
	a.m_code=b.m_code and a.m_code=c.m_code and b.s_code=c.s_code and a.m_name like @select_name order by a.m_code

end

if @status=3
begin

	select a.m_code,a.m_name,b.s_code,b.s_name,c.rate from
	test_info_main a,test_info_sub b, test_info_rate c where
	a.m_code=b.m_code and a.m_code=c.m_code and b.s_code=c.s_code and b.s_name like @select_name order by b.s_code

end

if @status=4
begin
		select u_id,u_name from  micropass where u_name like @select_name
end

if @status=5
begin
		select u_id,u_name from  micropass where u_id=@select_name
end

if @status=6
begin
        select doc_name from doctor_Info where refer_code=@select_name

end

if @status=7
begin
      select *from micropass where u_id=@select_name

end

if @status=8
begin
      select * from soft_bag where scr_no=@select_name

end

if @status=9
begin
      select * from emp_info where emp_id=@select_name

end

if @status=10
begin
      select * from Lv_Setup where leave_type=@select_name

end

if @status=11
begin
      select * from item_info where item_code=@select_name

end

if @status=12
begin
      select * from sup_info where sup_id=@select_name

end

if @status=13
begin


	select emp_id from emp_info where emp_name=@select_name

end

if @status=14
begin


	select * from stock_in where inv_no=@select_name

end

if @status=15
begin


	select * from item_info where item_name=@select_name

end


if @status=16
begin
      select sup_id,sup_name,sup_add=isnull(sup_add,''),
	sup_phone=isnull(sup_phone,'') from sup_info

end

if @status='17'
begin
      select * from font order by screen_name

end

if @status='18'
begin
select*from font where screen_name=@select_name
end

if @status='19'
begin
select vat_per=isnull(vat_per,0) from vat_per_setup
end

if @status='20'
begin
	select item_name from item_info where item_name like @select_name
end

if @status='21'
begin
	select emp_name from emp_info where emp_name like @select_name
end

if @status='22'
begin
      select * from emp_info where emp_name=@select_name

end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  PROCEDURE pro_name_SELECT1
@status int,
@pat_id int

AS


if @status=1
begin


	declare @refer_code as varchar(10)
	set @refer_code=(select refer_code from pat_info_main where pat_id=@pat_id)
	if @refer_code<>''
	begin
	
	select distinct a.pat_id,a.pat_name,a.sex,a.age,a.addr,a.phone,a.email,a.fax,
	a.refer_code,f.doc_name,
	a.uid,m.u_name,a.booth,a.dt1,b.delv_dt 
	from pat_info_main a,doctor_info f,micropass m,pat_info_sub1 b
	where a.refer_code=f.refer_code 
	and a.uid=m.u_id 
	and a.pat_id=b.pat_id 
	and a.pat_id=@pat_id

	end

	if @refer_code=''
	begin
	
	select distinct a.pat_id,a.pat_name,a.sex,a.age,a.addr,a.phone,a.email,a.fax,
	a.refer_code,f.doc_name,
	a.uid,m.u_name,a.booth,a.dt1,b.delv_dt 
	from pat_info_main a,doctor_info_new f,micropass m,pat_info_sub1 b
	where a.pat_id=f.pat_id 
	and a.uid=m.u_id 
	and a.pat_id=b.pat_id 
	and a.pat_id=@pat_id

	end




end

if @status=2
begin


	select b.m_code,b.s_code,t.s_name from pat_info_sub1 b,test_info_sub t where 
	b.m_code=t.m_code and b.s_code=t.s_code and b.pat_id=@pat_id
end



if @status=3
begin


	select * from pat_info_main where pat_id=@pat_id

end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE   PROCEDURE [pro_pass_entry] 

AS

insert into micropass (u_id,u_name,user_pass,uid,udt,cancel)

values('dsl','Daffodil Software Ltd.','44612123456','Default',getdate(),0)	--pass word dsl


insert into permit (code,u_id)values(16,'dsl')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.pro_security    Script Date: 08/11/2001 6:16:59 PM ******/
CREATE PROCEDURE [pro_security]
@carry as varchar(50),
@cur_user as varchar(20)

 AS
declare @Result as char(1)

set @Result= (select Case @carry
	
    when 'frmAdvance' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmAdvance') and u_id=@cur_user)

    when 'frmCompany_Info' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmCompany_Info') and u_id=@cur_user)

    when 'frmDoc_List' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmDoc_List') and u_id=@cur_user)

    when 'frmDoctor_Info' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmDoctor_Info') and u_id=@cur_user)

    when 'frmDoctor_Info_New' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmDoctor_Info_New') and u_id=@cur_user)

    when 'frmPatient_Info' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmPatient_Info') and u_id=@cur_user)

    when 'frmTest_Info' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmTest_Info') and u_id=@cur_user)

    when 'frmTest_List' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmTest_List') and u_id=@cur_user)

    when 'frmTest_Result' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmTest_Result') and u_id=@cur_user)

    when 'rBio_Chamical' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rBio_Chamical') and u_id=@cur_user)

    when 'rCancer' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rCancer') and u_id=@cur_user)

    when 'rDoc_New' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rDoc_New') and u_id=@cur_user)

    when 'rDoc_Pay' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rDoc_Pay') and u_id=@cur_user) END)

if @Result = 1 goto stop


set @Result= (select Case @carry

    when 'rEchocardiography' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rEchocardiography') and u_id=@cur_user)

    when 'rHepatitis' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rHepatitis') and u_id=@cur_user)

    when 'rHistopathology' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rHistopathology') and u_id=@cur_user)

    when 'rHormone' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rHormone') and u_id=@cur_user)

    when 'rImmunology' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rImmunology') and u_id=@cur_user)

    when 'rPaps' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rPaps') and u_id=@cur_user)

    when 'RptDoctor_Info' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='RptDoctor_Info') and u_id=@cur_user)

    when 'RptTest_Info' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='RptTest_Info') and u_id=@cur_user)END)


--if @Result <> null goto stop
if @Result =1 goto stop

set @Result= (select Case @carry

    when 'RptViewer' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='RptViewer') and u_id=@cur_user)

    when 'rStool' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rStool') and u_id=@cur_user)
    
    when 'rStool1' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rStool1') and u_id=@cur_user)

    when 'rUltrasonography' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rUltrasonography') and u_id=@cur_user)
    
    when 'rX_Ray' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='rX_Ray') and u_id=@cur_user)
    
    when 'frmCreate_User' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmCreate_User') and u_id=@cur_user)	
    
    when 'frmChange_Password' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmChange_Password') and u_id=@cur_user)
    
    when 'frmSoftware_Priviliege' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmSoftware_Priviliege') and u_id=@cur_user)
   
    when 'frmSoftware_Maintanance' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmSoftware_Maintanance') and u_id=@cur_user) END)

stop:

select Result=@Result













/*
set @Result= (select Case @carry
	
    when 'frmAdvance' then (select count(code) from permit
    where code=(select code from soft_bag where software='PRIME' and scr_no='frmAdvance') and u_id=@cur_user)

    when 'Form3' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form3') and u_id=@cur_user)

    when 'Form4' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form4') and u_id=@cur_user)

    when 'Form5' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form5') and u_id=@cur_user)

    when 'Form6' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form6') and u_id=@cur_user)

    when 'Form7' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form7') and u_id=@cur_user)

    when 'Form8' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form8') and u_id=@cur_user)

    when 'Form9' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form9') and u_id=@cur_user)

    when 'Form10' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form10') and u_id=@cur_user)

    when 'Form12' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form12') and u_id=@cur_user)

    when 'Form13' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form13') and u_id=@cur_user)

    when 'Form14' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form14') and u_id=@cur_user)

    when 'Form15' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form15') and u_id=@cur_user) END)

if @Result = 1 goto stop


set @Result= (select Case @carry

    when 'Form16' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form16') and u_id=@cur_user)

    when 'Form17' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form17') and u_id=@cur_user)

    when 'Form18' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form18') and u_id=@cur_user)

    when 'Form19' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form19') and u_id=@cur_user)

    when 'Form20' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form20') and u_id=@cur_user)

    when 'Form21' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form21') and u_id=@cur_user)

    when 'Form22' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form22') and u_id=@cur_user)

    when 'Form27' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form27') and u_id=@cur_user)END)

 --   when 'Form28' then (select count(code) from permit
   -- where code=(select code from soft_bag where software='Ayman PMIS' and scr_no='Form28') and u_id=@cur_user)

--    when 'Form29' then (select count(code) from permit
  --  where code=(select code from soft_bag where software='Ayman PMIS' and scr_no='Form29') and u_id=@cur_user) END)

--    when 'Form30' then (select count(code) from permit
--    where code=(select code from soft_bag where software='Ayman PMIS' and scr_no='Form30') and u_id=@cur_user) END)



--if @Result <> null goto stop
if @Result =1 goto stop

set @Result= (select Case @carry
	
   -- when 'Form31' then (select count(code) from permit
   -- where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form31') and u_id=@cur_user)

--    when 'Form32' then (select count(code) from permit
  --  where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form32') and u_id=@cur_user)

    when 'Form33' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form33') and u_id=@cur_user)

    when 'Form34' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form34') and u_id=@cur_user)
         when 'Form35' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form35') and u_id=@cur_user)

    when 'Form36' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form36') and u_id=@cur_user)
    
    when 'Form37' then (select count(code) from permit     where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form37') and u_id=@cur_user)

    when 'Form38' then (select count(code) from permit
    where code=(select code from soft_bag where software='Liberty PMIS' and scr_no='Form38') and u_id=@cur_user) END)

stop:
 select Result=@Result

*/














GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE      PROCEDURE pro_security_entry
AS
DELETE FROM SOFT_BAG

insert into soft_bag (software,scr_no,descript)values ('PRIME','frmChange_Password','For user Password Changing Screen')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmCommission_Per','for Doctor commission')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmCompany_Info','Only Company Information')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmCreate_User','Create New User')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmDoctor_Info','Doctors Information')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmEmp_Info','Employee Information')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmFont','Test Reports Font Changing Screen')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmItem_Info','Item Information Entry')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmLeave','Leave Information')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmLeave_Setup','Leave Setup')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmPatient_Info_VAT','Patient Information for VAT')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmPatient_Info','Patient Information')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmStock_IN','Item Purchase Information')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmStock_Out','Item Issue Information')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmSup_Info','Supplier Information')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmTest_Info','Test Information')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmTest_Result','Test Result Entry Screen')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmUser_Authority','Select User Permision')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmVAT_Setup','VAT percent SETUP for patient Test')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmPat_Info_Due','Patient Due Collection Screen')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmPay_Edit','Patient Payment Modification')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmDisc_Edit','Discount Modification')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rAdv_Coll','Advance or Due C0llection Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rBio_Chamical','Biochamical Examination Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rBody_Fluid','Body Fluid Examination Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rCT_SCAN','C.T. Scan Examination Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rDrug','DRUG TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rEchocardiography','ECHOCARDIOGRAPHY TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rEndoscopy','ENDOSCOPY TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rHaematology','HAEMATOLOGY TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rHepatitis','HEPATITIS TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rHistopathology','HISTOPATHOLOGY TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rHormone','HORMONE TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rImmunology','IMMUNOLOGY TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rMammography','MAMMOGRAPHY TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rMicrobiology','MICROBIOLOGY TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rPaps','PAPS Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rStool','STOOL TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rTumour_Marker','TUMOUR_MARKER Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rUltrasonogram','ULTRASONOGRAM Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rUrine1','URINE TEST Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rX_Ray','X-RAY Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rBooth_User_Info','Booth User Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rDaily_Statement','Monthly Statement Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rDoc_Due_Pat','Doctor Due Patient Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rDoc_New','New Doctor Information Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rDoc_Pay','Monthly Doctor Payment Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rLeave_Balance','Employee Leave Balance Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rPat_Info','Patient Test Information Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rPat_Type','Patient Type Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','RptDoctor_Info','Doctor Information Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','RptTest_Info','Test Information Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','rStock_Status','Stock Status Report')
insert into soft_bag (software,scr_no,descript)values ('PRIME','frmEdit_TestCode_Type','Modify Test Code and Group')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.rpt    Script Date: 16/10/2001 6:41:36 PM ******/
CREATE      PROCEDURE dbo.rpt

@status int,
@pat_id varchar(50),
@m_code varchar(2),
@s_code varchar(3)

AS
if @status=1


/*

IF (select COUNT(doc_name) from doctor_info where refer_code=(
	select refer_code from pat_info_main where pat_id=@pat_id))=0
BEGIN

select a.pat_id,a.m_code,a.s_code,a.field1,a.field2,a.field3,a.field3,a.field4,a.field5,a.field6,
	a.field7,a.field8,a.field9,a.field10,a.field11,a.field12,a.field13,
	a.field14,a.field15,a.uid,
        (select u_name from micropass where u_id=(select top 1 uid from report_all where pat_id=@pat_id)) usr_name,
	a.dt print_date,a.type,
	(select top 1 c.delv_dt from pat_info_sub1 c where c.pat_id=@pat_id) delv_date,
             b.pat_name,b.sex,b.age,b.addr,b.phone,b.fax,b.email,b.dt rec_date,doc_name =(
	select doc_name from doctor_info_new where pat_id=@pat_id)
	from report_all a,pat_info_main b
	where b.pat_id=a.pat_id and a.pat_id=@pat_id and a.m_code=@m_code and a.s_code=@s_code
	order by unique_id


END
ELSE
BEGIN

	select a.pat_id,a.m_code,a.s_code,a.field1,a.field2,a.field3,a.field3,a.field4,a.field5,a.field6,
	a.field7,a.field8,a.field9,a.field10,a.field11,a.field12,a.field13,
	a.field14,a.field15,a.uid,
        (select u_name from micropass where u_id=(select top 1 uid from report_all where pat_id=@pat_id)) usr_name,
	a.dt print_date,a.type,
	(select top 1 c.delv_dt from pat_info_sub1 c where c.pat_id=@pat_id) delv_date,
             b.pat_name,b.sex,b.age,b.addr,b.phone,b.fax,b.email,b.dt rec_date,doc_name =(
	select doc_name from doctor_info where refer_code=(
	select refer_code from pat_info_main where pat_id=@pat_id))
	from report_all a,pat_info_main b
	where b.pat_id=a.pat_id and a.pat_id=@pat_id and a.m_code=@m_code and a.s_code=@s_code
	order by unique_id

END
*/


IF (select COUNT(doc_name) from doctor_info where refer_code=(
	select refer_code from pat_info_main where pat_id=@pat_id))=0
BEGIN

select a.pat_id,a.m_code,a.s_code,a.field1,a.field2,a.field3,a.field3,a.field4,a.field5,a.field6,
	a.field7,a.field8,a.field9,a.field10,a.field11,a.field12,a.field13,
	a.field14,a.field15,a.uid,m.u_name usr_name,
	a.dt print_date,a.type,
	(select top 1 c.delv_dt from pat_info_sub1 c where c.pat_id=@pat_id) delv_date,
             b.pat_name,b.sex,b.age,b.addr,b.phone,b.fax,b.email,b.tmp_dt rec_date,doc_name =(
	select doc_name from doctor_info_new where pat_id=@pat_id),a.pat_id1
	from report_all a,pat_info_main b,micropass m
	where b.pat_id=a.pat_id and a.pat_id=@pat_id and a.m_code=@m_code
	and m.u_id=a.uid
	order by unique_id


END
ELSE
BEGIN

	select a.pat_id,a.m_code,a.s_code,a.field1,a.field2,a.field3,a.field3,a.field4,a.field5,a.field6,
	a.field7,a.field8,a.field9,a.field10,a.field11,a.field12,a.field13,
	a.field14,a.field15,a.uid,m.u_name usr_name,
	a.dt print_date,a.type,
	(select top 1 c.delv_dt from pat_info_sub1 c where c.pat_id=@pat_id) delv_date,
             b.pat_name,b.sex,b.age,b.addr,b.phone,b.fax,b.email,b.tmp_dt rec_date,doc_name =(
	select doc_name from doctor_info where refer_code=(
	select refer_code from pat_info_main where pat_id=@pat_id)),a.pat_id1
	from report_all a,pat_info_main b,micropass m
	where b.pat_id=a.pat_id and a.pat_id=@pat_id and a.m_code=@m_code
	and m.u_id=a.uid
	order by unique_id

END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE     PROC rptTest_State

@mode int,
@m_code varchar(2),
@st_date datetime,
@ed_date datetime

As
set nocount on

--set @st_date='2002-11-01 10:25:33.993'
--set @ed_date='2002-11-30 10:25:33.993'

--drop table #tmp1
--drop table #test_info_haema
--drop table #test_info_DRUG
--drop table #test_info_MAMMO
--drop table #test_info_CT
--drop table #test_info_OTHERS
--drop table #test_info_ULTRA
--drop table #test_info_X_RAY
--drop table #test_info_BODY
--drop table #test_info_STOOL
--drop table #test_info_URINE
--drop table #test_info_MICRO
--drop table #test_info_TUMOUR
--drop table #test_info_HOR
--drop table #test_info_HEPA
--drop table #test_info_IMMU
--drop table #test_info_BIO

if @mode=1

begin

---->>>FOR HAEMATOLOGY-->>>

	----******
	select m_code,s_code,s_name into #test_info_haema
	from test_info_sub where  m_code='01'
	--select*from #test_info_haema
	---****
	--COUNT 01-->>
	DECLARE @tot_haema int
	set @tot_haema=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='01'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_haema 
	into #tmp1 from pat_info_sub1 a,test_info_main b,#test_info_haema c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='01'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END HAEMATOLOGY--<<


---->>>FOR BIO-->>>



	----******
	select m_code,s_code,s_name into #test_info_BIO
	from test_info_sub where  m_code='02'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_BIO int
	set @tot_BIO=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='02'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	
	insert into #tmp1
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_BIO
	from pat_info_sub1 a,test_info_main b,#test_info_BIO c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='02'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END BIO--<<

---->>>FOR IMMU-->>>


	----******
	select m_code,s_code,s_name into #test_info_IMMU
	from test_info_sub where  m_code='03'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_IMMU int
	set @tot_IMMU=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='03'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_IMMU
	from pat_info_sub1 a,test_info_main b,#test_info_IMMU c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='03'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END IMMU--<<

---->>>FOR HEPA-->>>


	----******
	select m_code,s_code,s_name into #test_info_HEPA
	from test_info_sub where  m_code='04'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_HEPA int
	set @tot_HEPA=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='04'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_HEPA
	from pat_info_sub1 a,test_info_main b,#test_info_HEPA c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='04'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END HEPA--<<

---->>>FOR HOR-->>>


	----******
	select m_code,s_code,s_name into #test_info_HOR
	from test_info_sub where  m_code='05'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_HOR int
	set @tot_HOR=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='05'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_HOR
	from pat_info_sub1 a,test_info_main b,#test_info_HOR c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='05'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END HOR--<<

---->>>FOR TUMOUR-->>>


	----******
	select m_code,s_code,s_name into #test_info_TUMOUR
	from test_info_sub where  m_code='06'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_TUMOUR int
	set @tot_TUMOUR=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='06'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_TUMOUR
	from pat_info_sub1 a,test_info_main b,#test_info_TUMOUR c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='06'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END TUMOUR--<<

---->>>FOR MICRO-->>>


	----******
	select m_code,s_code,s_name into #test_info_MICRO
	from test_info_sub where  m_code='07'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_MICRO int
	set @tot_MICRO=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='07'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_MICRO
	from pat_info_sub1 a,test_info_main b,#test_info_MICRO c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='07'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END MICRO--<<


---->>>FOR URINE-->>>

	----******
	select m_code,s_code,s_name into #test_info_URINE
	from test_info_sub where  m_code='08'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_URINE int
	set @tot_URINE=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='08'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_URINE
	from pat_info_sub1 a,test_info_main b,#test_info_URINE c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='08'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END URINE--<<


---->>>FOR STOOL-->>>


	----******
	select m_code,s_code,s_name into #test_info_STOOL
	from test_info_sub where  m_code='09'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_STOOL int
	set @tot_STOOL=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='09'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_STOOL
	from pat_info_sub1 a,test_info_main b,#test_info_STOOL c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='09'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END STOOL--<<


---->>>FOR BODY-->>>


	----******
	select m_code,s_code,s_name into #test_info_BODY
	from test_info_sub where  m_code='10'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_BODY int
	set @tot_BODY=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='10'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_BODY
	from pat_info_sub1 a,test_info_main b,#test_info_BODY c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='10'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END BODY--<<


---->>>FOR X_RAY-->>>


	----******
	select m_code,s_code,s_name into #test_info_X_RAY
	from test_info_sub where  m_code='11'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_X_RAY int
	set @tot_X_RAY=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='11'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_X_RAY
	from pat_info_sub1 a,test_info_main b,#test_info_X_RAY c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='11'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END X_RAY--<<

---->>>FOR ULTRA-->>>


	----******
	select m_code,s_code,s_name into #test_info_ULTRA
	from test_info_sub where  m_code='12'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_ULTRA int
	set @tot_ULTRA=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='12'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_ULTRA
	from pat_info_sub1 a,test_info_main b,#test_info_ULTRA c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='12'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END ULTRA--<<

---->>>FOR OTHERS-->>>


	----******
	select m_code,s_code,s_name into #test_info_OTHERS
	from test_info_sub where  m_code='13'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_OTHERS int
	set @tot_OTHERS=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='13'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_OTHERS
	from pat_info_sub1 a,test_info_main b,#test_info_OTHERS c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='13'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END OTHERS--<<


---->>>FOR CT-->>>


	----******
	select m_code,s_code,s_name into #test_info_CT
	from test_info_sub where  m_code='14'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_CT int
	set @tot_CT=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='14'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_CT
	from pat_info_sub1 a,test_info_main b,#test_info_CT c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='14'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END CT--<<


---->>>FOR MAMMO-->>>


	----******
	select m_code,s_code,s_name into #test_info_MAMMO
	from test_info_sub where  m_code='15'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_MAMMO int
	set @tot_MAMMO=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='15'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_MAMMO
	from pat_info_sub1 a,test_info_main b,#test_info_MAMMO c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='15'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp1
---<<<END OTHERS--<<

---->>>FOR DRUG-->>>


	----******
	select m_code,s_code,s_name into #test_info_DRUG
	from test_info_sub where  m_code='16'
	--select*from #test_info_sub
	---****
	--COUNT 01-->>
	DECLARE @tot_DRUG int
	set @tot_DRUG=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code='16'
	and a.dt1 between @st_date and @ed_date )
	--<<---
	insert into #tmp1 
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_DRUG
	from pat_info_sub1 a,test_info_main b,#test_info_DRUG c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code='16'
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
	--select*from #tmp1
	select m_code,m_name,s_code,s_name,tot_sub=count(s_code),Tot_test from #tmp1
	group by m_code,m_name,s_code,s_name,Tot_test

---<<<END DRUG--<<

end

if @mode=2
begin

	----******
	select m_code,s_code,s_name into #test_info_haema1
	from test_info_sub where  m_code=@m_code
	--select*from #test_info_haema1
	---****
	--COUNT 01-->>
	DECLARE @tot_haema1 int
	set @tot_haema1=(select Tot_test=count(a.m_code) from pat_info_sub1 a
	where a.m_code=@m_code
	and a.dt1 between @st_date and @ed_date )
	--<<---
	
	select a.m_code,b.m_name,c.s_code,c.s_name,Tot_test=@tot_haema1 
	into #tmp11 from pat_info_sub1 a,test_info_main b,#test_info_haema1 c
	where a.m_code=b.m_code 
	and c.s_code=a.s_code
	and a.m_code=@m_code
	and a.dt1 between @st_date and @ed_date 
	--group by a.m_code,b.m_name
	
--	select*from #tmp11
	select m_code,m_name,s_code,s_name,tot_sub=count(s_code),Tot_test from #tmp11
	group by m_code,m_name,s_code,s_name,Tot_test

end

set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.sp_found    Script Date: 26/09/2001 9:06:52 AM ******/
CREATE PROCEDURE [sp_found] 

@m_code as varchar(2),
@s_code as varchar(3)

AS

if @m_code = '' and @s_code <>''
begin
	select found=0
	goto stop
end

if @m_code <> '' and @s_code =''
begin
	if (select count(m_code) from test_info_rate where  m_code=@m_code)  > 0
		select found='Y'
	else
		select found='N'
	goto stop
end

if @m_code <> '' and @s_code <>''
begin
	if (select count(*) from test_info_rate where  m_code=@m_code and s_code=@s_code)  > 0
		select a.s_name,b.rate,a.type from test_info_sub a , test_info_rate b where a.m_code=b.m_code
		and a.s_code=b.s_code and a.m_code=@m_code and a.s_code=@s_code
	else
		select found='N'
end


stop:







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.test_Info_SELECT    Script Date: 26/09/2001 9:06:53 AM ******/
CREATE PROCEDURE test_Info_SELECT

@Status int,
@m_code varchar(2)

 AS

if @Status=1

begin
	select m_code from test_info_main where m_code=@m_code
end

if @Status=2

begin
	select max(pat_id) as pat_id from pat_info_main where booth=@m_code
end










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE test_Result_Select7

@mode int,
@test_name varchar(100),
@m_code varchar(2),
@s_code varchar(2)
As

if @mode=1
begin
select * from test_result where m_code=@m_code and s_code=@s_code and test_name=@test_name
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






CREATE    PROCEDURE test_result_SELECT

@Status int,
@test_name varchar(2000),
@type varchar(4)

 AS

if @Status=1

begin

select test_name=isnull(test_name,''),test_result=isnull(test_result,''),unit=isnull(unit,''),
ref_range=isnull(ref_range,''),type=isnull(type,''),others=isnull(others,''),others1=isnull(others1,''),
m_code=isnull(m_code,''),s_code=isnull(s_code,''),unique_id from
test_result where test_name=@test_name and type=@type

end

if @Status=2

begin

--select * from test_result where test_result=@test_name and type=@type

select test_name=isnull(test_name,''),test_result=isnull(test_result,''),unit=isnull(unit,''),
ref_range=isnull(ref_range,''),type=isnull(type,''),others=isnull(others,''),others1=isnull(others1,''),
m_code=isnull(m_code,''),s_code=isnull(s_code,''),unique_id from
test_result where test_name=@test_name and type=@type

end


if @Status=3

begin

select test_name=isnull(test_name,''),test_result=isnull(test_result,''),unit=isnull(unit,''),
ref_range=isnull(ref_range,''),type=isnull(type,''),others=isnull(others,''),others1=isnull(others1,''),
m_code=isnull(m_code,''),s_code=isnull(s_code,''),unique_id from
test_result where test_result=@test_name and type=@type

end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.test_result_SELECT1    Script Date: 15/10/2001 6:26:42 PM ******/
CREATE PROCEDURE test_result_SELECT1

@Status int,
@type varchar(4)

 AS

if @Status=1

begin

select distinct test_name  from test_result where type=@type

end

if @Status=2

begin

select distinct unit  from test_result where type=@type

end

if @Status=3

begin

select distinct test_result  from test_result where type=@type

end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.test_result_SELECT2    Script Date: 15/10/2001 6:26:42 PM ******/
CREATE    PROCEDURE test_result_SELECT14



as

select *  from test_result 
where type='12' 



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







/****** Object:  Stored Procedure dbo.test_result_SELECT2    Script Date: 15/10/2001 6:26:42 PM ******/
CREATE   PROCEDURE test_result_SELECT2



as

select *  from test_result 
where type<>'01' 
order by type








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO





/****** Object:  Stored Procedure dbo.test_result_SELECT3    Script Date: 20/01/2002 10:27:03 AM ******/
CREATE PROCEDURE test_result_SELECT3

@Status int,
@unique_id int

 AS

if @Status=1

begin

     select * from test_result where unique_id=@unique_id

end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.test_result_SELECT4    Script Date: 28/01/2002 06:13:53 PM ******/
CREATE PROCEDURE test_result_SELECT4

@Status int,
@type varchar(100),
@test_name varchar(2000)

 AS

if @Status=1

begin

select test_result from test_result where type=@type and  test_name=@test_name

end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/****** Object:  Stored Procedure dbo.test_result_SELECT5    Script Date: 28/01/2002 06:13:53 PM ******/
CREATE PROCEDURE test_result_SELECT5

@Status int,
@test_name varchar(2000),
@test_result varchar(2000),
@type varchar(4)

 AS

if @Status=1

begin
	select * from test_result where test_name=@test_name 
	and test_result=@test_result and type=@type
end

if @Status=2

begin
select ref_range from test_result where test_name=@test_name and test_result=@test_result and type=@type
end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE test_result_SELECT6

@mode int

as

if @mode=1
begin
		select *  from test_result where type='01' order by type
end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE    PROCEDURE test_result_SELECT9


@m_code varchar(4)

as

select *  from test_result
where m_code=@m_code
order by type




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

