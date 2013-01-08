namespace HouseCostCalculation
{
    public class Address
    {
        public string fullAddress(string type)
        {
            string fullAddress = "";
            string buildNum = null;
            if (type == "квартира")
            {
                if (building != "")
                {
                    buildNum = ", корп. " + building;
                }

                fullAddress = " квартира №" + room + " " + town + ", " + street + ", " + house + buildNum;
            }
            else if (type == "домовладение")
            {
                if (building != "")
                {
                    buildNum = ", корп. " + building;
                }

                fullAddress = " домовладение " + town + ", " + street + ", " + house + buildNum;
            }
            else if (type == "земельный участок")
            {
                fullAddress = " земельный участок " + town + ", " + street + ", " + house + buildNum;
            }
            return fullAddress;
        }

        private string town;
        private string street;
        private string region;
        private string building;
        private int room;
        private int house;

        public string Town
        {
            get
            {
                return town;
            }
            set
            {
                town = value;
            }
        }

        public string Street
        {
            get
            {
                return street;
            }
            set
            {
                street = value;
            }
        }

        public int House
        {
            get
            {
                return house;
            }
            set
            {
                house = value;
            }
        }

        public string Building
        {
            get
            {
                return building;
            }
            set
            {
                building = value;
            }
        }

        public int Room
        {
            get
            {
                return room;
            }
            set
            {
                room = value;
            }
        }

        public string Region
        {
            get
            {
                return region;
            }
            set
            {
                region = value;
            }
        }
    }
}