using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HouseCostCalculation;


namespace HouseCostCalculation
{
    public partial class ownerForm : Form
    {
        public ownerForm()
        {
            InitializeComponent();
        }

        public ownerForm(List<Owner> owners)
        {
            InitializeComponent();
            
            foreach(HouseCostCalculation.Owner owner in owners)
            {
                ownerList.Rows.Add(owner.ownerName, owner.ownerInit, owner.ownerSurname, owner.address, owner.passportSerial, owner.passNum, owner.OVD, owner.passDate);
            }

            Show();
            Activate();
        }
    }
}
