using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelWriteRead
{
    /// <summary>
    /// excel操作类
    /// </summary>
    public class ExcelOperator
    {
        /// <summary>
        /// 读取Excel
        /// </summary>
        /// <returns></returns>
        public static DataTable ReadExcel()
        {
            var dt = new DataTable();
            //读取文件的路径
            string path = ConfigurationManager.AppSettings["ExcelPath"];
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + "Extended Properties=Excel 8.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "select * from [sheet1$]";
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strExcel, strConn);
            myCommand.Fill(dt);
            conn.Close();
            return dt; 
        }

       
        /// <summary>
        /// 将数据导入到CSV
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static bool WriteCSV(DataTable dt)
        {
            //写入文件的路径
            string fileNmae = ConfigurationManager.AppSettings["CSVPath"];
            bool isSucess = false;
            string con = "";
            foreach (DataColumn dc in dt.Columns)
            {
                con += dc.ColumnName + ",";
            }
            con = con.TrimEnd(',') + Environment.NewLine;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    con += dt.Rows[i][j].ToString().Replace("\n", " ").Replace("\r\n", " ").Replace(",", "，") + ",";
                }
                con = con.TrimEnd(',') + Environment.NewLine;
            }
            try
            {
                FileStream fs = new FileStream(fileNmae, FileMode.Create);
                byte[] b = Encoding.GetEncoding("utf-8").GetBytes(con);
                fs.Write(b, 0, b.Length);
                fs.Close();
                isSucess = true;
            }
            catch
            {
                throw;
            }
            return isSucess;
        }

        //将数据写入已存在Excel
        public static void writeExcel(string result, string filepath)
        {
            //1.创建Applicaton对象
            //Microsoft.Office.Interop.Excel.Application xApp = new

            //Microsoft.Office.Interop.Excel.Application();

            ////2.得到workbook对象，打开已有的文件
            //Microsoft.Office.Interop.Excel.Workbook xBook = xApp.Workbooks.Open(filepath,
            //                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            //                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            //                      Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            ////3.指定要操作的Sheet
            //Microsoft.Office.Interop.Excel.Worksheet xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[1];

            ////在第一列的左边插入一列  1:第一列
            ////xlShiftToRight:向右移动单元格   xlShiftDown:向下移动单元格
            ////Range Columns = (Range)xSheet.Columns[1, System.Type.Missing];
            ////Columns.Insert(XlInsertShiftDirection.xlShiftToRight, Type.Missing);

            ////4.向相应对位置写入相应的数据
            //xSheet.Cells[Column(列)][Row(行)] = result;

            ////5.保存保存WorkBook
            //xBook.Save();
            ////6.从内存中关闭Excel对象

            //xSheet = null;
            //xBook.Close();
            //xBook = null;
            ////关闭EXCEL的提示框
            //xApp.DisplayAlerts = false;
            ////Excel从内存中退出
            //xApp.Quit();
            //xApp = null;
        }
    }
}
