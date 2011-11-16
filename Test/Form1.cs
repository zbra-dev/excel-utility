using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ExcelUtility;
using System.Xml;
using System.Text.RegularExpressions;

namespace Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (ExcelFile file = ExcelFile.Open(@"D:\temp\Sheet3.xlsx"))
            {
                IWorksheet sheet1 = file.OpenWorksheet("Paosdpoasdp");
                Column column = sheet1.GetColumn("E");
                double realPosition = sheet1.GetColumnPosition(column.ColumnIndex);
                Row row = sheet1.GetRow(2);
                double rowRealPosition = sheet1.GetRowPosition(row.Index);

                //column.Width = 4.57;

                //Numbers
                Cell h1 = sheet1.GetCell("H1"); //Cell in use.
                h1.Value = "40210";
                Cell c1 = sheet1.GetCell("C1");
                c1.Value = "3443";
                
                //Text
                Cell b1 = sheet1.GetCell("B1");
                b1.Value = "Atwood Falcon"; //Existing text - index = 7;
                Cell a3 = sheet1.GetCell("A3");
                a3.Value = "New Text for A3"; //new Text
                
                /*Shape shape = sheet1.DrawShape(0, 0, 100, 100);
                shape.ForeColor = Color.Black;
                shape.MarginLeft = 10;
                shape.MarginRight = 10;
                shape.MarginTop = 10;
                shape.MarginBottom = 10;
                shape.Text = "12345";
                 * */
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int groups = 3;
            int months = 12;
            using (ExcelFile file = ExcelFile.Open(@"D:\temp\Sheet3.xlsx"))
            {
                var sheet = file.OpenWorksheet("AC1");
                for (int i = 0; i < groups; ++i)
                {
                    var column = sheet.GetColumn(((char)'A' + i).ToString());
                    column.Width = 100;
                    var cell = sheet.GetCell(column.Name + "1");
                    cell.Value = "Group " + i;
                }
                for (int i = 0; i < months; ++i)
                {
                    var column = sheet.GetColumn(((char)'A' + i + groups).ToString());
                    column.Width = 50;
                    //column.Color = i % 2 == 0 ? "Gray" : "White";
                    var cell = sheet.GetCell(column.Name + "1");
                    cell.Value = "Month " + i;
                }
                // Row 1
                var row = sheet.GetRow(2);
                row.Height = 100;
                for (int i = 0; i < groups; ++i)
                {
                    var columnName = ((char)'A' + i).ToString();
                    var cell = sheet.GetCell(columnName + row.Index);
                    cell.Value = "Value " + i;
                }
                var start = sheet.GetColumn(((char)'A' + groups).ToString());
                var shape1 = sheet.Drawing.DrawShape(start, row, 0, 0, 250, 50);
                shape1.Text = "Shape1";
                var shape2 = sheet.Drawing.DrawShape(start, row, 50, 25, 250, 50);
                shape2.Text = "Shape2";
            }
        }
    }
}
