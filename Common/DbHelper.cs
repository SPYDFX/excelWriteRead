using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace excelWriteRead
{
    /// <summary>
    /// 数据库连接类
    /// </summary>
    public class DbHelper
    {
        /// <summary>
        /// 获取数据库连接信息
        /// </summary>
        private static string strcon
        {
            get
            {
                return ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
            }
        }

        /// <summary>
        /// 查询数据库获取指定表
        /// </summary>
        /// <param name="tbName"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static DataTable GetTable(string tbName)
        {
            SqlConnection connection = new SqlConnection(strcon);
            connection.Open();
            string cmdStr = string.Format("select * from {0} ", tbName);
            SqlDataAdapter sqlda = new SqlDataAdapter(cmdStr, connection);
            DataTable dt = new DataTable();
            sqlda.Fill(dt);
            connection.Close();
            return dt;
        }
    }
}
