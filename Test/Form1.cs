using System;
using System.Windows.Forms;
using ExcelUtility;

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
            using (ExcelFile file = ExcelFile.Open(@"D:\Book1.xlsx"))
            {
                IWorksheet sheet2 = file.OpenWorksheet("asasda");
                IWorksheet sheet1 = file.OpenWorksheet("Sheet1");
                var cell = sheet1.GetCell("A4");


                //IWorksheet sheet1 = file.OpenWorksheet("Sheet2");

                //Column a = sheet1.CreateColumnBetween(8, 8);


                //Column column = sheet1.GetColumn("F");
                //double realPosition = sheet1.GetColumnPosition(column.ColumnIndex);
                //Row row = sheet1.GetRow(12);
                //double rowRealPosition = sheet1.GetRowPosition(row.Index);

                //column.Width = 4.57;
                /*
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
                */

                /*var shape = sheet1.Drawings.DrawShape(sheet1.GetColumn("D"), sheet1.GetRow(4), 0, 0, 40, 80);
                shape.Text = "Shape1";
                shape.ForeColor = Color.Black;
                shape.MarginLeft = 10;
                shape.MarginRight = 10;
                shape.MarginTop = 10;
                shape.MarginBottom = 10;
                shape.Text = "12345";*/
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int groups = 3;
            int months = 12;
            using (ExcelFile file = ExcelFile.Open(@"D:\temp\Book1.xlsx"))
            {
                var sheet = file.OpenWorksheet("Sheet1");
                for (int i = 0; i < groups; ++i)
                {
                    var column = sheet.GetColumn(i);
                    column.Width = 15;
                    var cell = sheet.GetCell(column.Name + "1");
                    cell.StringValue = "Group " + i;
                }
                for (int i = 0; i < months; ++i)
                {
                    var column = sheet.GetColumn(i + groups);
                    column.Width = 7;
                    //column.Color = i % 2 == 0 ? "Gray" : "White";
                    var cell = sheet.GetCell(column.Name + "1");
                    cell.StringValue = "Month " + i;
                }
                // Row 1
                var row = sheet.GetRow(2);
                //row.Height = 100;
                for (int i = 0; i < groups; ++i)
                {
                    var column = sheet.GetColumn(i);
                    var cell = sheet.GetCell(column.Name + row.Index);
                    cell.StringValue = "Value " + i;
                }
                var start = sheet.GetColumn(groups);
                //var shape1 = sheet.Drawings.DrawShape(start, row, 0, 0, 250, 50);
                //shape1.Text = "Shape1";
                //var shape2 = sheet.Drawings.DrawShape(start, row, 50, 25, 250, 50);
                //shape2.Text = "Shape2";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (ExcelFile file = ExcelFile.Open(@"D:\temp\Book1.xlsx"))
            {
                var sheet = file.OpenWorksheet("Sheet1");
                var shape = sheet.DrawShape(2, 4, 2, 4, 4, 4, 4, 4);
                shape.Text = "123";
            }
        }
    }
}
