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

namespace Data_Compiler
{
    public partial class Form2 : Form
    {

        StreamReader reader = null;
        private MainForm main = null;

        public Form2(Form mainFormCall)
        {
            main = mainFormCall as MainForm;
            InitializeComponent();
        }
       
        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            main.SetButtonEnable(true);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            String fileName = Path.Combine(Directory.GetCurrentDirectory(), "\\Models.txt");
            try
            {
                if ((reader = new StreamReader(fileName)) != null)
                {
                    Task.Factory.StartNew(ConvertData);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
            }
        }

        private void ConvertData()
        {
            dataGridView1.Rows.Add();
        }
    }
}
