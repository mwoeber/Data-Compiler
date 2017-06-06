using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Data_Compiler
{
    public partial class Form2 : Form
    {
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
    }
}
