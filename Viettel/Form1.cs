//using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Viettel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        [DllImport("user32.dll")]

        static extern void mouse_event(int dwFlags, int dx, int dy, int dwdata, int dwextrainfo);

        public enum mouseeventflags
        {
            // Bấm nút
            LeftDown = 2,
            // Thả nút
            LeftUp = 4,
        }

        public void leftclick(Point p)
        {
            mouse_event((int)(mouseeventflags.LeftDown), p.X, p.Y, 0, 0);
            mouse_event((int)(mouseeventflags.LeftUp), p.X, p.Y, 0, 0);
        }

        bool stop = true;
        
        private void button1_Click(object sender, EventArgs e)
        {
            stop =  (stop) ? false : true;
            timer1.Interval = (int)numericUpDown1.Value;
            timer1.Enabled = true;

            if (!stop) timer1.Start();
            if(stop) timer1.Stop();
            
        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            
            // Kiểm tra nếu không có hàng nào trong DataGridView
            if (dataGridView1.Rows.Count == 0)
                return;

            // Lấy chỉ mục của hàng hiện tại (index của hàng đầu tiên)
            int currentRowIndex = dataGridView1.FirstDisplayedCell.RowIndex;
  

            // Kiểm tra xem chỉ mục hàng hiện tại có lớn hơn hoặc bằng số hàng trong DataGridView không
            if (currentRowIndex >= dataGridView1.Rows.Count)
                return;

            // Lấy giá trị từ ô trong hàng hiện tại
            string cellValue = dataGridView1.Rows[currentRowIndex].Cells[5].Value?.ToString(); // Lưu ý sử dụng ?. để kiểm tra null trước khi gọi ToString()

            // Kiểm tra xem giá trị có rỗng không
            if (string.IsNullOrEmpty(cellValue))
                return;

            // Gửi giá trị của ô đến ứng dụng đích
            SendKeys.SendWait(cellValue);

            // Gửi phím Enter
            SendKeys.SendWait("{ENTER}");

            // Dịch chuyển sang hàng tiếp theo trong DataGridView
            dataGridView1.FirstDisplayedCell = dataGridView1.Rows[currentRowIndex + 1].Cells[0];
        }

        
        private void btnAdd_Click(object sender, EventArgs e)
        {
            OpenFileDialog dilg = new OpenFileDialog();
            dilg.Filter = "Excel Sheet(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            if (dilg.ShowDialog() == DialogResult.OK)
            {
                string filepath = dilg.FileName;
                textBox1.Text = filepath;

                LoadDataFromExceltoDataGridView(filepath, ".xlsx", "yes");
            }
        }


        public void LoadDataFromExceltoDataGridView(string fpath, string ext, string hdr)
        {
            string con = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
            con = string.Format(con, fpath, "yes");
            OleDbConnection excelconnection = new OleDbConnection(con);
            excelconnection.Open();
            DataTable dtexcel = excelconnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string excelsheet = dtexcel.Rows[0]["TABLE_NAME"].ToString();
            OleDbCommand com = new OleDbCommand("Select * from [" + excelsheet + "]", excelconnection);
            OleDbDataAdapter oda = new OleDbDataAdapter(com);
            DataTable dt = new DataTable();
            oda.Fill(dt);
            excelconnection.Close();

            dataGridView1.DataSource = dt;
        }


    }
}
