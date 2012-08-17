using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WindowsFormsApplication1;
namespace HouseCostCalculation
{
    public partial class startForm : Form
    {
        public mainForm mainForm;
    
        public startForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.Text != null)
            {
                Visible = false;
                this.mainForm = new mainForm(listBox1.Text, banksList.Text);
                mainForm.Show();
            }
        }
    }
}
