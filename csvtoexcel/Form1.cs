using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace csvtoexcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog()
            {
                Filter = ".CSV Dosyaları|*.csv",
                Multiselect = false,
                Title = "Select CSV file to export:"
            };

            DialogResult f = o.ShowDialog();
            if (f != DialogResult.Cancel)
            {
                textBox1.Text = o.FileName;
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("you must select a csv file");
                return;
            }

            DataTable data = new DataTable();
            bool first_move = true;
            int count = 0;
            string[] rows = File.ReadAllLines(textBox1.Text, Encoding.GetEncoding("windows-1254")); // for turkish chars, you can use UTF8 (Encoding.UTF8)
            foreach (string row  in rows)
            {
                // create columns
                if (first_move)
                {
                    count = row.Split(',').Length;
                    for (int i = 0; i < 20; i++) // if you have spesific column count variable
                    {
                        data.Columns.Add("Sütun#" + (i).ToString());
                    }

                    first_move = false;
                }

                string[] split = row.Split(',');
                List<string[]> lst = new List<string[]>();

                data.Rows.Add();
                DataRow dr = data.Rows[data.Rows.Count - 1];
                
                dr.ItemArray = split;
            }

            dataGridView1.DataSource = data;
            list_created = true;
        }

        bool list_created = false;
        private void Button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("you must select a csv file");
                return;
            }

            if (!list_created)
            {
                MessageBox.Show("data list not created!");
                return;
            }

            try
            {
                // creating Excel Application  
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                // creating new WorkBook within Excel application  
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                // creating new Excelsheet in workbook  
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                // see the excel sheet behind the program  
                app.Visible = true;
                // get the reference of first sheet. By default its name is Sheet1.  
                // store its reference to worksheet  
                worksheet = workbook.Sheets["Sayfa1"];
                worksheet = workbook.ActiveSheet;
                // changing the name of active sheet  
                worksheet.Name = "export";
                // storing header part in Excel  
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                // storing Each row and column value to excel sheet  
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
                // save the application  
                workbook.SaveAs(Application.StartupPath + "\\output.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application  
                app.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Process interrupted. Please don't click anywhere while list creating");
            }

            MessageBox.Show("File created.\r\n\r\n" + Application.StartupPath + "\\output.xlsx");
        }
    }
}
