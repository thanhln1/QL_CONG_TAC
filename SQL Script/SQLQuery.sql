select * from TMP_WORKING  where  ID='200' and MA_KHACH_HANG='';
select * from MT_WORKING  where ID  = '85'

truncate table HIS_WORKING

INSERT INTO HIS_WORKING SELECT * FROM TMP_WORKING

select distinct(MA_KHACH_HANG) from HIS_WORKING  where MA_KHACH_HANG!='' and cast (WORKING_DAY as date)  between @from and @to;

select cast (WORKING_DAY as date) from TMP_WORKING


Select * into TMP_HOP_DONG  from  MT_HOP_DONG; 

select * from HIS_HOP_DONG

Update TMP_HOP_DONG set CHI_PHI_THUC_DA_CHI = CHI_PHI_THUC_DA_CHI+123;

select * from HIS_WORKING order by ID

select * from TMP_WORKING

select * from MT_WORKING where DATEPART(DW,WORKING_DAY) != '1' and MA_NHAN_VIEN = 'NVA'

SELECT * from MT_NHAN_VIEN where MA_NHAN_VIEN = (select MA_NHAN_VIEN from TMP_WORKING where ID = '1')

SELECT DATENAME(DW, GETDATE())

UPDATE MT_WORKING a set a.MARK = (select b.ID from MT_HOP_DONG b where b.MA_KHACH_HANG = a.MA_KHACH_HANG);

select b.ID from MT_HOP_DONG b where b.MA_KHACH_HANG = 'ABC'
truncate table TMP_HOP_DONG
select * from TMP_HOP_DONG
SELECT DON_GIA FROM MT_DON_GIA WHERE DIA_CHI = (SELECT TINH FROM MT_HOP_DONG WHERE MA_KHACH_HANG='ABC');