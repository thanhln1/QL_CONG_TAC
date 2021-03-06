﻿using Dapper;
using DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
   public class HIS_SCHEDUAL_DAO
    {

        COMMON dao = new COMMON();


        public void SaveSchedual( HIS_SCHEDUAL shedual, int month, int year )
        {
            using (IDbConnection cnn = new System.Data.SqlClient.SqlConnection(dao.ConnectionString("Default")))
            {

                StringBuilder sql = new StringBuilder();
                sql.Append("insert into MT_SCHEDUAL ");
                sql.Append("(MA_NHAN_VIEN, THANG, NAM,");
                sql.Append(" TUAN1_THU2, TUAN1_THU3, TUAN1_THU4, TUAN1_THU5, TUAN1_THU6, TUAN1_THU7, TUAN1_CN,");
                sql.Append(" TUAN2_THU2, TUAN2_THU3, TUAN2_THU4, TUAN2_THU5, TUAN2_THU6, TUAN2_THU7, TUAN2_CN,");
                sql.Append(" TUAN3_THU2, TUAN3_THU3, TUAN3_THU4, TUAN3_THU5, TUAN3_THU6, TUAN3_THU7, TUAN3_CN,");
                sql.Append(" TUAN4_THU2, TUAN4_THU3, TUAN4_THU4, TUAN4_THU5, TUAN4_THU6, TUAN4_THU7, TUAN4_CN) ");
                sql.Append(" values "); 
                sql.Append("(@MA_NHAN_VIEN, @THANG, @NAM,");
                sql.Append(" @TUAN1_THU2, @TUAN1_THU3, @TUAN1_THU4, @TUAN1_THU5, @TUAN1_THU6, @TUAN1_THU7, @TUAN1_CN,");
                sql.Append(" @TUAN2_THU2, @TUAN2_THU3, @TUAN2_THU4, @TUAN2_THU5, @TUAN2_THU6, @TUAN2_THU7, @TUAN2_CN,");
                sql.Append(" @TUAN3_THU2, @TUAN3_THU3, @TUAN3_THU4, @TUAN3_THU5, @TUAN3_THU6, @TUAN3_THU7, @TUAN3_CN,");
                sql.Append(" @TUAN4_THU2, @TUAN4_THU3, @TUAN4_THU4, @TUAN4_THU5, @TUAN4_THU6, @TUAN4_THU7, @TUAN4_CN) ");
                cnn.Execute(sql.ToString(), shedual);
            }
        }      
    }
}
