using Padeg;

namespace HouseCostCalculation
{
    public class Owner
    {
        private string ownerAddress;
        private string ownerOVD;
        private string initials;
        private string firstName;
        private string lastName;
        private string passSerial;
        private string passRegDate;
        private string passNumber;
        private string ownerPhone;
        public string ownerFullName;
        public string ownerFullNameR;
        public string ownerFullNameD;
        public string ownerFullNameV;
        public string ownerFullNameT;
        public string ownerFullNameP;

        public Owner(string address1, string OVD1, string ownerInit1, string ownerName1, string ownerSurname1, string passDate1, string passNum1, string passportSerial1, string phone1)
        {
            address = address1;

            OVD = OVD1;
            ownerInit = ownerInit1;
            ownerName = ownerName1;
            ownerSurname = ownerSurname1;
            passportSerial = passportSerial1;
            passNum = passNum1;
            passDate = passDate1;
            phone = phone1;
            ownerPadeg();
            ownerFullName = ownerSurname + " " + ownerName + " " + ownerInit;
        }

        public string ownerName
        {
            get
            {
                return firstName;
            }
            set
            {
                firstName = value;
            }
        }

        public string ownerSurname
        {
            get
            {
                return lastName;
            }
            set
            {
                lastName = value;
            }
        }

        public string ownerInit
        {
            get
            {
                return initials;
            }
            set
            {
                initials = value;
            }
        }

        public string phone
        {
            get
            {
                return ownerPhone;
            }
            set
            {
                ownerPhone = value;
            }
        }

        public string address
        {
            get
            {
                return ownerAddress;
            }
            set
            {
                ownerAddress = value;
            }
        }

        public string passNum
        {
            get
            {
                return passNumber;
            }
            set
            {
                passNumber = value;
            }
        }

        public string passDate
        {
            get
            {
                return passRegDate;
            }
            set
            {
                passRegDate = value;
            }
        }

        public string passportSerial
        {
            get
            {
                return passSerial;
            }
            set
            {
                passSerial = value;
            }
        }

        public string OVD
        {
            get
            {
                return ownerOVD;
            }
            set
            {
                ownerOVD = value;
            }
        }

        public void ownerPadeg()
        {
            Declension padeg = new Declension();
            int sex = padeg.GetSex(ownerInit);
            string cSex;
            if (sex == 1)
            {
                cSex = "м";
            }
            else
            {
                cSex = "ж";
            }

            ownerFullNameR = padeg.GetFIOPadeg(ownerSurname, ownerName, ownerInit, cSex, 2);
            ownerFullNameD = padeg.GetFIOPadeg(ownerSurname, ownerName, ownerInit, cSex, 3);
            ownerFullNameV = padeg.GetFIOPadeg(ownerSurname, ownerName, ownerInit, cSex, 4);
            ownerFullNameT = padeg.GetFIOPadeg(ownerSurname, ownerName, ownerInit, cSex, 5);
            ownerFullNameP = padeg.GetFIOPadeg(ownerSurname, ownerName, ownerInit, cSex, 6);
        }
    }
}