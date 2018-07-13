using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelWriteRead
{
    public class UserOp
    {
        public DataTable GetUserInfo()
        {
           return ExcelOperator.ReadExcel();
        }

        public bool WriteCSV()
        {
            bool isSucess = false;
            var dtResult = new DataTable();
            dtResult.Columns.Add("用户编号", typeof(string));
            dtResult.Columns.Add("用户姓名", typeof(string));
            dtResult.Columns.Add("用户电话号码", typeof(string));
            dtResult.Columns.Add("用户性别", typeof(string));
            var dt= DbHelper.GetTable("userInfo");
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    DataRow drow = dtResult.NewRow();
                    drow["用户编号"] = row[0];
                    drow["用户姓名"] = row[1];
                    drow["用户电话号码"] = row[2];
                    string sex = "无";
                    if(row[3]!=null)
                    {
                        if(int.Parse(row[3].ToString())==1)
                        {
                            sex = "男";
                        }
                        else
                        {
                            sex = "女";
                        }
                    }
                    drow["用户性别"] = sex;
                    dtResult.Rows.Add(drow);
                }
            }
            isSucess= ExcelOperator.WriteCSV(dtResult);
            return isSucess;
        }
    }
}
