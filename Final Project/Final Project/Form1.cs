using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Final_Project
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string FileStr = "D:\\Test";

            Excel.Application Excel_APP1 = new Excel.Application();
            Excel.Workbook Excel_WB1 = Excel_APP1.Workbooks.Add();
            Excel.Worksheet Excel_WS1 = new Excel.Worksheet();
            Excel_WS1 = Excel_WB1.Worksheets[1];
            Excel_WS1.Name = "004";

            Excel_APP1.Cells[1, 1] = "名稱";
            Excel_APP1.Cells[1, 2] = "數量";

            Excel_APP1.Cells[2, 1] = "奶茶";
            Excel_APP1.Cells[2, 2] = "1";

            Excel_APP1.Cells[3, 1] = "蘋果";
            Excel_APP1.Cells[3, 2] = "1";

            Excel_APP1.Cells[4, 1] = "鳳梨";
            Excel_APP1.Cells[4, 2] = "1";

            Excel_WB1.SaveAs(FileStr);

            Excel_WS1 = null;
            Excel_WB1.Close();
            Excel_WB1 = null;
            Excel_APP1.Quit();
            Excel_APP1 = null;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string FileStr = "D:\\Test";

            Excel.Application Excel_APP1 = new Excel.Application();
            Excel.Workbook Excel_WB1 = Excel_APP1.Workbooks.Open(FileStr);
            Excel.Worksheet Excel_WS1 = new Excel.Worksheet();
            Excel_WS1 = Excel_WB1.Worksheets[1];
            Excel_WS1.Name = "004";

            Excel_APP1.Cells[5, 1] = "香蕉";
            Excel_APP1.Cells[5, 2] = "3";

            Excel_WB1.Save();

            Excel_WS1 = null;
            Excel_WB1.Close();
            Excel_WB1 = null;
            Excel_APP1.Quit();
            Excel_APP1 = null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string FileStr = "D:\\Test";

            Excel.Application Excel_APP1 = new Excel.Application();
            Excel.Workbook Excel_WB1 = Excel_APP1.Workbooks.Open(FileStr);
            Excel.Worksheet Excel_WS1 = new Excel.Worksheet();
            Excel_WS1 = Excel_WB1.Worksheets[1];
            Excel_WS1.Name = "004";

            Excel_APP1.Cells[5, 1].Delete();
            Excel_APP1.Cells[5, 1].Delete();

            Excel_WB1.Save();

            Excel_WS1 = null;
            Excel_WB1.Close();
            Excel_WB1 = null;
            Excel_APP1.Quit();
            Excel_APP1 = null;
        }
    }
}
