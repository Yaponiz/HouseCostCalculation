using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Padeg;
using WindowsFormsApplication1;
using HouseCostCalculation.Properties;

namespace HouseCostCalculation
{
    public partial class Padeg : Form
    {
        public int t;
        public Padeg()
        {
            InitializeComponent();
        }

        public Padeg(string fullNameR, string fullNameD, string fullNameV, string fullNameT, string fullNameP, int oc)
        {
            InitializeComponent();
            t = oc;
            Declension padeg = new Declension();
            string firstName = null;
            string lastName = null;
            string init = null;
            padeg.SeparateFIO(fullNameR, ref lastName,  ref firstName, ref init);
            textBox1.Text = lastName;
            textBox2.Text = firstName;
            textBox3.Text = init;

            padeg.SeparateFIO(fullNameD, ref lastName, ref firstName, ref init);
            textBox4.Text = lastName;
            textBox5.Text = firstName;
            textBox6.Text = init;

            this.Show();
            this.Activate();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            mainForm f = new mainForm();
            
            f.fullNameRSet(textBox1.Text + " " + textBox2.Text + " " + textBox3.Text, t);
            f.fullNameDSet(textBox4.Text + " " + textBox5.Text + " " + textBox6.Text, t);
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
