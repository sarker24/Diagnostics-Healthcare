if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[F1]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[F1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[F2]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[F2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[F3]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[F3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[F4]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[F4]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[F5]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[F5]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[F6]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[F6]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[F7]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[F7]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[getSalesManPerformance]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[getSalesManPerformance]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE FUNCTION F1 (@refer_code varchar(10),@StDate datetime,@EdDate datetime)
RETURNS @pat_info_main TABLE
   (
    Pat_id     int,
    Pat_Name   varchar(80),
    refer_code varchar(10),	
    vat_amt money,
    disc money
   )
AS
BEGIN

insert into @pat_info_main

select a.pat_id,a.pat_name,a.refer_code,a.vat_amt,b.disc 
from pat_info_main a,pat_info_sub3 b
where refer_code=@refer_code and a.pat_id=b.pat_id 
and a.dt between @StDate and @EdDate
RETURN
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




CREATE  FUNCTION F2 (@refer_code varchar(10),@StDate datetime,@EdDate datetime)
RETURNS @pat_info_sub TABLE
   (
    Pat_id     int,
    test_rate money
   )
AS
BEGIN

insert into @pat_info_sub
select b.pat_id,test_rate=sum(b.test_rate) 
from pat_info_sub1 b where b.pat_id=(select pat_id from dbo.F1(@refer_code,@StDate,@EdDate))
group by b.pat_id
RETURN
END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




CREATE  FUNCTION F3 (@refer_code varchar(10),@StDate datetime,@EdDate datetime)
RETURNS @temp2 TABLE
   (
    Pat_id     int,
    vat_amt money,	
    disc money,	
    test_rate money
   )
AS
BEGIN
insert into @temp2
select m.pat_id,m.vat_amt,m.disc,b.test_rate 
from F2(@refer_code,@StDate,@EdDate) b,F1(@refer_code,@StDate,@EdDate) m
where m.pat_id=b.pat_id
RETURN
END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE FUNCTION F4 (@refer_code varchar(10),@StDate datetime,@EdDate datetime)
RETURNS @temp3 TABLE
   (
    Pat_id     int,
    adv money,	
    collect_fee money
   )
AS
BEGIN
insert into @temp3
select c.pat_id,adv=sum(c.adv),collect_fee=sum(c.collect_fee) 
from pat_info_sub2 c,F1(@refer_code,@StDate,@EdDate) m 
where m.pat_id=c.pat_id
group by c.pat_id
RETURN
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




CREATE  FUNCTION F5 (@refer_code varchar(10),@StDate datetime,@EdDate datetime)
RETURNS @temp4 TABLE
   (
	Pat_id     int,
	vat_amt money,
	test_rate money,
	disc money,	
	adv money,	
	collect_fee money,
	Due money
   )
AS
BEGIN
insert into @temp4
select a.pat_id,a.vat_amt,a.test_rate,a.disc,b.adv,b.collect_fee,
Due=(a.test_rate+a.vat_amt+b.collect_fee-a.disc-b.adv) 
from F3(@refer_code,@StDate,@EdDate) a,F4(@refer_code,@StDate,@EdDate) b
where a.pat_id=b.pat_id

RETURN
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE FUNCTION F6 (@refer_code varchar(10),@StDate datetime,@EdDate datetime)
RETURNS @temp5 TABLE
   (
    Pat_id     int
   )
AS
BEGIN
insert into @temp5
select pat_id from F5(@refer_code,@StDate,@EdDate) where due>0

RETURN
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE  FUNCTION F7 (@refer_code varchar(10),@StDate datetime,@EdDate datetime)
RETURNS @temp6 TABLE
   (
    Pat_id     int,
    pat_name Varchar(60),
    addr Varchar(100),
    phone Varchar(25),
    refer_code Varchar(10),
    doc_name Varchar(60)
   )
AS
BEGIN
insert into @temp6
select a.pat_id,a.pat_name,a.addr,a.phone,a.refer_code,d.doc_name 
from pat_info_main a,F6(@refer_code,@StDate,@EdDate) b,doctor_info d
where a.pat_id=b.pat_id and a.refer_code=d.refer_code

RETURN
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



-- SELECT * FROM getSalesManPerformance ('52','01 feb 2006','28 feb 2006')

CREATE    FUNCTION getSalesManPerformance(
	@SalesManId VARCHAR (2),
	@DateFrom DATETIME,
	@DateTo DATETIME)

RETURNS @TmpTable  Table(
	DoctorId VARCHAR (10),
	DoctorName VARCHAR (100),
	PATH MONEY,
	SPATH MONEY,
	HISTO MONEY,
	XRAY MONEY,
	ECG MONEY,
	USG MONEY,
	ECHO MONEY,
	ENDO MONEY,
	DOPL MONEY,
	Total MONEY)
AS
BEGIN
	DECLARE @DoctorId VARCHAR (10),
			@DoctorName VARCHAR (100),
			@PATH MONEY,
			@SPATH MONEY,
			@HISTO MONEY,
			@XRAY MONEY,
			@ECG MONEY,
			@USG MONEY,
			@ECHO MONEY,
			@ENDO MONEY,
			@DOPL MONEY,
			@Total MONEY

	DECLARE MyCursor CURSOR FOR
	SELECT refer_code, doc_name FROM Doctor_Info WHERE (EmpId = @SalesManId)

	OPEN MyCursor
		FETCH NEXT FROM MyCursor INTO @DoctorId, @DoctorName
		WHILE @@FETCH_STATUS = 0 
		BEGIN
			SELECT @PATH = ISNULL(SUM(Pat_Info_Sub1.test_rate),0) FROM Pat_Info_Sub1 INNER JOIN Pat_Info_Main ON 
			Pat_Info_Sub1.pat_id = Pat_Info_Main.pat_id WHERE (Pat_Info_Main.refer_code = @DoctorId) AND (Pat_Info_Sub1.type = 'PATH')
			AND (Pat_Info_Sub1.tmp_dt BETWEEN @DateFrom AND @DateTo)

			SELECT @SPATH = ISNULL(SUM(Pat_Info_Sub1.test_rate),0) FROM Pat_Info_Sub1 INNER JOIN Pat_Info_Main ON 
			Pat_Info_Sub1.pat_id = Pat_Info_Main.pat_id WHERE (Pat_Info_Main.refer_code = @DoctorId) AND (Pat_Info_Sub1.type = 'SPATH')
			AND (Pat_Info_Sub1.tmp_dt BETWEEN @DateFrom AND @DateTo)

			SELECT @HISTO = ISNULL(SUM(Pat_Info_Sub1.test_rate),0) FROM Pat_Info_Sub1 INNER JOIN Pat_Info_Main ON 
			Pat_Info_Sub1.pat_id = Pat_Info_Main.pat_id WHERE (Pat_Info_Main.refer_code = @DoctorId) AND (Pat_Info_Sub1.type = 'HISTO')
			AND (Pat_Info_Sub1.tmp_dt BETWEEN @DateFrom AND @DateTo)

			SELECT @XRAY = ISNULL(SUM(Pat_Info_Sub1.test_rate),0) FROM Pat_Info_Sub1 INNER JOIN Pat_Info_Main ON 
			Pat_Info_Sub1.pat_id = Pat_Info_Main.pat_id WHERE (Pat_Info_Main.refer_code = @DoctorId) AND (Pat_Info_Sub1.type = 'X-RAY')
			AND (Pat_Info_Sub1.tmp_dt BETWEEN @DateFrom AND @DateTo)

			SELECT @ECG = ISNULL(SUM(Pat_Info_Sub1.test_rate),0) FROM Pat_Info_Sub1 INNER JOIN Pat_Info_Main ON 
			Pat_Info_Sub1.pat_id = Pat_Info_Main.pat_id WHERE (Pat_Info_Main.refer_code = @DoctorId) AND (Pat_Info_Sub1.type = 'ECG')
			AND (Pat_Info_Sub1.tmp_dt BETWEEN @DateFrom AND @DateTo)

			SELECT @USG = ISNULL(SUM(Pat_Info_Sub1.test_rate),0) FROM Pat_Info_Sub1 INNER JOIN Pat_Info_Main ON 
			Pat_Info_Sub1.pat_id = Pat_Info_Main.pat_id WHERE (Pat_Info_Main.refer_code = @DoctorId) AND (Pat_Info_Sub1.type = 'USG')
			AND (Pat_Info_Sub1.tmp_dt BETWEEN @DateFrom AND @DateTo)

			SELECT @ECHO = ISNULL(SUM(Pat_Info_Sub1.test_rate),0) FROM Pat_Info_Sub1 INNER JOIN Pat_Info_Main ON 
			Pat_Info_Sub1.pat_id = Pat_Info_Main.pat_id WHERE (Pat_Info_Main.refer_code = @DoctorId) AND (Pat_Info_Sub1.type = 'ECHO')
			AND (Pat_Info_Sub1.tmp_dt BETWEEN @DateFrom AND @DateTo)

			SELECT @ENDO = ISNULL(SUM(Pat_Info_Sub1.test_rate),0) FROM Pat_Info_Sub1 INNER JOIN Pat_Info_Main ON 
			Pat_Info_Sub1.pat_id = Pat_Info_Main.pat_id WHERE (Pat_Info_Main.refer_code = @DoctorId) AND (Pat_Info_Sub1.type = 'ENDO')
			AND (Pat_Info_Sub1.tmp_dt BETWEEN @DateFrom AND @DateTo)

			SELECT @DOPL = ISNULL(SUM(Pat_Info_Sub1.test_rate),0) FROM Pat_Info_Sub1 INNER JOIN Pat_Info_Main ON 
			Pat_Info_Sub1.pat_id = Pat_Info_Main.pat_id WHERE (Pat_Info_Main.refer_code = @DoctorId) AND (Pat_Info_Sub1.type = 'DOPL')
			AND (Pat_Info_Sub1.tmp_dt BETWEEN @DateFrom AND @DateTo)
			
			SET @Total = @PATH + @SPATH + @HISTO + @XRAY + @ECG + @USG + @ECHO + @ENDO + @DOPL
			INSERT INTO @TmpTable VALUES (@DoctorId, @DoctorName, @PATH, @SPATH, @HISTO, @XRAY, @ECG, @USG, @ECHO, @ENDO, @DOPL, @Total)
		FETCH NEXT FROM MyCursor INTO @DoctorId, @DoctorName
		END
	CLOSE MyCursor
	DEALLOCATE MyCursor
RETURN
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

