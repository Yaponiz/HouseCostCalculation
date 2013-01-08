using System.Collections.Generic;
using System.Windows.Forms;

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

            foreach (HouseCostCalculation.Owner owner in owners)
            {
                ownerList.Rows.Add(owner.ownerName, owner.ownerInit, owner.ownerSurname, owner.address, owner.passportSerial, owner.passNum, owner.OVD, owner.passDate);
            }

            Show();
            Activate();
        }
    }
}