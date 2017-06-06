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
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace Data_Compiler
{
    public partial class MainForm : Form
    {
        String fileName = null;
        StreamReader reader = null;
        int __progress = 0;

        public MainForm()
        {
            InitializeComponent();
        }

        private delegate void Delegate();

        private void button1_Click(object sender, EventArgs e)
        {
            
            
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;


            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            { 
                try
                {
                    fileName = openFileDialog1.FileName.ToString();
                    if ((reader = new StreamReader(fileName)) != null)
                    {
                       Task.Factory.StartNew( ConvertData );
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }

        }

        private void ConvertData()
        {            

            SetButtonEnable(false);

            String fileName_ = fileName.Substring(0, fileName.Length - 4);
            fileName_ = fileName_.Replace(".", string.Empty);
            String[] fileNameArray = fileName_.Split(Path.DirectorySeparatorChar);
            String path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            String Path_ = path + Path.DirectorySeparatorChar + fileNameArray[fileNameArray.Length - 1].Trim() + ".xlsx";

            _Application docExcel = new Microsoft.Office.Interop.Excel.Application();

            docExcel.Visible = false;
            docExcel.DisplayAlerts = false;

            _Workbook workbooksExcel = docExcel.Workbooks.Add();
            _Worksheet worksheetExcel = (_Worksheet)workbooksExcel.ActiveSheet;

            var rng_ = ((Range)worksheetExcel.Range[worksheetExcel.Cells[1, 1], worksheetExcel.Cells[1, 17]]);
            rng_.Value = new string[,] { { "Source Code", "", "", "", "", "", "", "ORD DATE", "COPIES", "GROSS VALUE", "NET VALUE", "GROSS RE PER COP", "NET REV. PER COPY", "GROSS REVENUE", "NET REVENUE", "TOTAL" } };

            int count = 1;

            Stream baseStream = reader.BaseStream;
            long length = baseStream.Length;

            while (reader.Peek() >= 0)
            {
                String line = reader.ReadLine().Trim();
                String[] array = line.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                if (array.Length > 0 && array[0].Any(char.IsDigit) && array[0].Trim() != "")
                {
                    if (array[0].Length == 6 || (array[0].Length == 5 && array[0].Substring(0, 1) == "Z"))
                    {
                        count++;

                        var rng = ((Range)worksheetExcel.Range[worksheetExcel.Cells[count, 1], worksheetExcel.Cells[count, 17]]);
                        rng.Value = new string[,] { { array[0], array[0].Substring(0, 1), array[0].Substring(1, 1), array[0].Substring(2, 1), array[0].Substring(3, 1), array[0].Substring(4, 1), (array[0].Length == 6 ? array[0].Substring(5, 1) : ""), array[1], array[2], array[3], array[4], array[5], array[6], array[7], array[8], array[9], array[10] } };

                    }
                }

                long progress = (baseStream.Position * 100 / length);
                __progress = (int)(progress) ;

                if (progressBar1.InvokeRequired)
                {
                    var myDelegate = new Delegate(UpdateProgressBar);
                    progressBar1.Invoke(myDelegate);
                }
                else
                {
                    UpdateProgressBar();
                }
            }

            workbooksExcel.SaveAs(Path_, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, System.Reflection.Missing.Value, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, Type.Missing, Type.Missing, Type.Missing);

            workbooksExcel.Close(false, Type.Missing, Type.Missing);
            docExcel.Application.DisplayAlerts = true;
            docExcel.Application.Quit();

            SetButtonEnable(true);
        }

        public void SetButtonEnable(bool text)
        {
            if (InvokeRequired)
            {
                Invoke((Action<bool>)SetButtonEnable, text);
                return;
            }
            button1.Enabled = text;
        }

        private void UpdateProgressBar()
        {
            progressBar1.Value = __progress;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void prefrencesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetButtonEnable(false);
            Form2 preference = new Form2(this);
            preference.Show();
        }
    }
}
