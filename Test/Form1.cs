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
                //Column column = sheet1.GetColumn("E");
                //column.Width = 4.57;

                //Numbers
                Cell e1 = sheet1.GetCell("E1"); //Cell in use.
                e1.Value = "40210";
                Cell c1 = sheet1.GetCell("C1");
                c1.Value = "3443";
                
                //Text
                Cell b1 = sheet1.GetCell("B1");
                b1.Value = "Atwood Falcon"; //Existing text - index = 7;
                Cell d1 = sheet1.GetCell("D1");
                d1.Value = "New Text"; //new Text

                /*
                Shape shape = sheet1.DrawShape(0, 0, 100, 100);
                shape.ForeColor = Color.Black;
                shape.MarginLeft = 10;
                shape.MarginRight = 10;
                shape.MarginTop = 10;
                shape.MarginBottom = 10;
                shape.Text = "12345";
                */
            }
        }
    }
}
