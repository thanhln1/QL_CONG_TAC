select * from TMP_WORKING  where  ID='200' and MA_KHACH_HANG='';
select * from MT_WORKING  order by MA_NHAN_VIEN

truncate table MA_NHAN_VIEN;
truncate table MT_WORKING;
truncate table MT_HOP_DONG;


INSERT INTO HIS_WORKING SELECT * FROM TMP_WORKING

select distinct(MA_KHACH_HANG) from HIS_WORKING  where MA_KHACH_HANG!='' and cast (WORKING_DAY as date)  between @from and @to;

select cast (WORKING_DAY as date) from TMP_WORKING


Select * into TMP_HOP_DONG  from  MT_HOP_DONG; 

select * from MT_HOP_DONG

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

-- tạo bảng bằng câu select
Select * into HIS_HOP_DONG  from  MT_HOP_DONG
 
select * from HIS_HOP_DONG
select * from MT_WORKING order by MA_NHAN_VIEN
select * from HIS_WORKING order by MA_NHAN_VIEN

SELECT * FROM HIS_WORKING  WHERE ID IN (SELECT ID FROM MT_WORKING)
TRUNCATE TABLE MT_NHAN_VIEN

UPDATE MT_HOP_DONG
SET
    MT_HOP_DONG.CHI_PHI_THUC_DA_CHI = b.CHI_PHI_THUC_DA_CHI
FROM
    MT_HOP_DONG a
INNER JOIN
    TMP_HOP_DONG b
ON 
    a.ID = b.ID;

DELETE FROM HIS_WORKING WHERE ID IN(SELECT ID FROM TMP_WORKING)


-- Thanh kotex
select * from MT_NHAN_VIEN where MA_NHAN_VIEN in(select distinct a.MA_NHAN_VIEN from HIS_WORKING a where  cast (a.WORKING_DAY as date) BETWEEN  '2019-06-14 00:00:00.000' and '2019-07-17 00:00:00.000'  and a.MA_KHACH_HANG = 'ABC')

SELECT DATEDIFF(day, min(a. WORKING_DAY),  MAX(a.WORKING_DAY)) as DAY  from HIS_WORKING a where  cast (a.WORKING_DAY as date) BETWEEN  '2019-06-14 00:00:00.000' and '2019-07-17 00:00:00.000'  and a.MA_KHACH_HANG = 'ABC'

select * from MT_HOP_DONG where CHI_PHI_THUC_DA_CHI < TONG_CHI_PHI_MUC_TOI_DA order by newid();

SELECT  cast(min(a. WORKING_DAY) as Date)   from HIS_WORKING a where  cast (a.WORKING_DAY as date) BETWEEN  '2019-06-14 00:00:00.000' and '2019-07-17 00:00:00.000'  and a.MA_KHACH_HANG = 'ABC'

select max(c.COUNT_DAY) from (select count(*) as COUNT_DAY, MA_NHAN_VIEN from HIS_WORKING a where  cast (a.WORKING_DAY as date) BETWEEN  '2019-06-14 00:00:00.000' and '2019-07-17 00:00:00.000'  and a.MA_KHACH_HANG = 'ILM' group by MA_NHAN_VIEN) as C

select * from HIS_WORKING a where  cast (a.WORKING_DAY as date) BETWEEN  '2019-06-14 00:00:00.000' and '2019-07-17 00:00:00.000'  and a.MA_KHACH_HANG = 'ABC'

select count(*) as COUNT_DAY, MA_NHAN_VIEN from HIS_WORKING a where  cast (a.WORKING_DAY as date) BETWEEN  '2019-06-14 00:00:00.000' and '2019-07-17 00:00:00.000'  and a.MA_KHACH_HANG = 'ILM' group by MA_NHAN_VIEN order by MA_NHAN_VIEN

select * from TMP_WORKING where ID='394'

INSERT INTO TMP_HOP_DONG SELECT * FROM MT_HOP_DONG

UPDATE TMP_WORKING set MA_KHACH_HANG= 'CTTVINAWACO' WHERE ID = 394