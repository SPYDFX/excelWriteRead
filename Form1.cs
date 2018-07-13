using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excelWriteRead
{
    public partial class Form1 : Form
    {
        UserOp op = new UserOp();
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            dgvUserInfo.DataSource = op.GetUserInfo();
        }

        private void btnWrite_Click(object sender, EventArgs e)
        {
            if (op.WriteCSV())
            {
                MessageBox.Show("操作成功");
            }
            else
            {
                MessageBox.Show("操作失败");
            }
        }
    }
}
