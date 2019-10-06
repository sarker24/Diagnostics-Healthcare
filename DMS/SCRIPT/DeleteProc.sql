
--Pat_Info_Main
DELETE FROM Pat_Info_Main WHERE (SUBSTRING(pat_id1, 3, 2) NOT IN ('05', '06'))

--Pat_Info_Sub1
DELETE FROM Pat_Info_Sub1 WHERE (YEAR(tmp_dt) NOT IN ('2005','2006'))

--Pat_Info_Sub2
DELETE FROM Pat_Info_Sub2 WHERE (YEAR(tmp_dt) NOT IN ('2005','2006'))

--Pat_Info_Sub3
DELETE FROM Pat_Info_Sub3 WHERE (YEAR(tmp_dt) NOT IN ('2005','2006'))

