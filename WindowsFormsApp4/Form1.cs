using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Load += Form1_Load;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SpreadsheetControl spreadsheetControl = new SpreadsheetControl();
            string path = Path.Combine(Application.StartupPath, "견적내역서_입력폼.xlsx");
            spreadsheetControl.LoadDocument(path);
            spreadsheetControl.BeginUpdate();
            IWorkbook workbook = spreadsheetControl.Document;
            string sheetName = $"TEST_TEST";
            Worksheet worksheet = workbook.Worksheets["Sample"];
         
            spreadsheetControl.SaveDocument();

            workbook.Worksheets[sheetName].CopyFrom(worksheet);

            //workbook.Worksheets.Add();
            spreadsheetControl.SaveDocument();
            spreadsheetControl.EndUpdate();
            spreadsheetControl.Dispose();
        }
    }
}
