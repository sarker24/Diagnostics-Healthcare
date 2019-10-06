SELECT     TOP 100 PERCENT Pat_Info_Main.pat_id, Pat_Info_Main.pat_id1, Pat_Info_Main.tmp_dt, Pat_Info_Sub1.tmp_dt AS Expr1, 
                      Pat_Info_Sub2.tmp_dt AS Expr2, Pat_Info_Sub3.tmp_dt AS Expr3
FROM         Pat_Info_Main INNER JOIN
                      Pat_Info_Sub1 ON Pat_Info_Main.pat_id = Pat_Info_Sub1.pat_id INNER JOIN
                      Pat_Info_Sub2 ON Pat_Info_Main.pat_id = Pat_Info_Sub2.pat_id INNER JOIN
                      Pat_Info_Sub3 ON Pat_Info_Main.pat_id = Pat_Info_Sub3.pat_id
ORDER BY Pat_Info_Main.tmp_dt