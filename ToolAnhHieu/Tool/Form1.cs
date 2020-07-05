using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Random rd=new Random();
            string path = @"F:\Visual studio project\C# co ban\Tool\123.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(path);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];
          
            double add = 0.1;
           
                double Lp = Double.Parse(textBox1.Text.ToString().Trim());
                double Up = Double.Parse(textBox2.Text.ToString().Trim());
                for (int j = 2; j < 10; j++)
                {
                ReRun:
                    for (int i = 0; i < 10; i++)
                    {
                        ws.Cells[i + 8, j] = Math.Round(rd.NextDouble() * (Up - Lp) + Lp, 2).ToString();
                    }

                if (double.Parse(ws.Cells[7, j].Value.ToString()) < 1.5)
                    {
                        Lp += add;
                        Up -= add;
                        goto ReRun;
                    }
                }


            wb.Save();
            wb.Close();
            richTextBox1.Text = "ok";

        }
    }
}
