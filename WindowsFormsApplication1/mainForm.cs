using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using HouseCostCalculation;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Padeg;
using RSDN;
using WMPLib;
using Application = Microsoft.Office.Interop.Word.Application;
using DataTable = System.Data.DataTable;
using Shape = Microsoft.Office.Interop.Word.Shape;

namespace WindowsFormsApplication1
{
    [Serializable]
    public partial class mainForm : Form
    {
        public static Bank bank;
        public static string rooms1;
        public static string roomsT;
        public static string customerFullNameR;
        public static string customerFullNameD;
        public static string customerFullNameV;
        public static string customerFullNameT;
        public static string customerFullNameP;
        public static string customerFamiliyR, customerNameR, customerSurnameR;
        public static string ownerFullNameR;
        public static string ownerFullNameD;
        public static string ownerFullNameV;
        public static string ownerFullNameT;
        public static string ownerFullNameP;
        public static double cost1, cost2, cost3;
        public static double m1, m2, m3;
        public static double cost_m1, cost_m2, cost_m3;
        public static double cor_torg, cor_torg2, cor_torg3;
        public static double cor_cost1, cor_cost2, cor_cost3;
        public static string date1, date2, date3;
        public static double cor1, cor2, cor3;
        public static double cor_cost11, cor_cost21, cor_cost31;
        public static double cor_place1, cor_place2, cor_place3;
        public static double cor_cost12, cor_cost22, cor_cost32;
        public static double cor_type1, cor_type2, cor_type3;
        public static double cor_cost13, cor_cost23, cor_cost33;
        public static double cor_date1, cor_date2, cor_date3;
        public static double cor_cost14, cor_cost24, cor_cost34;
        public static double cor_floor1, cor_floor2, cor_floor3;
        public static double cor_cost15, cor_cost25, cor_cost35;
        public static double cor_m1, cor_m2, cor_m3;
        public static double cor_cost16, cor_cost26, cor_cost36;
        public static double cor_b1, cor_b2, cor_b3;
        public static double cor_cost17, cor_cost27, cor_cost37;
        public static double cor_height1, cor_height2, cor_height3;
        public static double cor_cost18, cor_cost28, cor_cost38;
        public static double cor_class1, cor_class2, cor_class3;
        public static double cor_cost19, cor_cost29, cor_cost39;
        public static double cor_phone1, cor_phone2, cor_phone3;
        public static double cor_cost110, cor_cost210, cor_cost310;
        public static double cor_com1, cor_com2, cor_com3;
        public static double cor_cost111, cor_cost211, cor_cost311;
        public static double cor_t1, cor_t2, cor_t3;
        public static double cor_cost112, cor_cost212, cor_cost312;
        public static double cor_lift1, cor_lift2, cor_lift3;
        public static double cor_cost113, cor_cost213, cor_cost313;
        public static int cost_count1 = 0, cost_count2 = 0, cost_count3 = 0;
        public static double cost_cor_koef1, cost_cor_koef2, cost_cor_koef3;
        public static double cor_cost_final1, cor_cost_final2, cor_cost_final3;
        public static double final_cost_m;
        public static double m_final;
        public static double finalCost;
        public static double finalCostRounded;
        public static double finalDirtCost;
        public static double likvidCost;
        public static double likvidCostDirt;
        public static string costStr = "0";
        public static string uvaj;
        public static string houseTypeAdds;
        public static string kadastr;
        public static string docType;
        public static string houseType1, houseType2;
        public static string docTypeT;

        public object Missing;
        public List<Owner> owners = new List<Owner>();
        public string roomsN;
        public string roomsX;
        public Application wdApp;

        public mainForm()
        {
            InitializeComponent();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void updateTables()
        {
            try
            {
                string ht1 = "";
                if (houseType.Text == "Кирпичный")
                {
                    ht1 = "Кирпичные";
                }

                if (houseType.Text == "Панельный")
                {
                    ht1 = "Панельные";
                }

                if (houseType.Text == "Монолитный")
                {
                    ht1 = "Монолитные";
                }

                string str;

                if (floors.Value == 1)
                {
                    str = "-но";
                }
                else if (floors.Value == 2)
                {
                    str = "-ух";
                }
                else if (floors.Value == 3)
                {
                    str = "-ех";
                }
                else if (floors.Value == 4)
                {
                    str = "-ех";
                }
                else if (floors.Value == 7)
                {
                    str = "-и";
                }
                else if (floors.Value == 8)
                {
                    str = "-и";
                }
                else
                {
                    str = "-ти";
                }

                if (floors.Value < 6)
                {
                    lift.SelectedIndex = 1;
                }
                else if (floors.Value > 6)
                {
                    lift.SelectedIndex = 0;
                }


                switch (docTypeT)
                {
                    case "Квартира":
                        {
                            objectDataGrid.Rows[43].Cells[1].Value = floor.Value.ToString() + " этаж";
                            analogsGrid.Rows[6].Cells[1].Value = floor.Value.ToString().ToLower();

                            //analogsGrid.Rows[6].Cells[2].Value = floor.Value.ToString().ToLower();
                            //analogsGrid.Rows[6].Cells[3].Value = floor.Value.ToString().ToLower();
                            //analogsGrid.Rows[6].Cells[4].Value = floor.Value.ToString().ToLower();
                            objectDataGrid.Rows[18].Cells[1].Value = ht1;
                            analogsGrid.Rows[2].Cells[1].Value = street.Text;
                            analogsGrid.Rows[2].Cells[2].Value = street.Text;
                            analogsGrid.Rows[2].Cells[3].Value = street.Text;
                            analogsGrid.Rows[2].Cells[4].Value = street.Text;

                            objectDataGrid.Rows[30].Cells[1].Value = lift.Text;
                            analogsGrid.Rows[14].Cells[1].Value = lift.Text.ToLower();
                            analogsGrid.Rows[14].Cells[2].Value = lift.Text.ToLower();
                            analogsGrid.Rows[14].Cells[3].Value = lift.Text.ToLower();
                            analogsGrid.Rows[14].Cells[4].Value = lift.Text.ToLower();
                            objectDataGrid.Rows[1].Cells[1].Value = fullAddress();
                            objectDataGrid.Rows[20].Cells[1].Value = floors.Value.ToString() + str;
                            analogsGrid.Rows[5].Cells[1].Value = floors.Value.ToString().ToLower();

                            objectDataGrid.Rows[54].Cells[0].Value =
                            "Общая площадь квартиры, согласно правоустанавливающим документам (" + registrationDoc.Text +
                            "), в кв.м.";

                            objectDataGrid.Rows[18].Cells[1].Value = houseType.Text;
                            objectDataGrid.Rows[21].Cells[1].Value = ht1;
                            objectDataGrid.Rows[22].Cells[1].Value = ht1;
                            analogsGrid.Rows[3].Cells[1].Value = houseType.Text.ToLower();
                            analogsGrid.Rows[3].Cells[2].Value = houseType.Text.ToLower();
                            analogsGrid.Rows[3].Cells[3].Value = houseType.Text.ToLower();
                            analogsGrid.Rows[3].Cells[4].Value = houseType.Text.ToLower();


                        }
                        break;

                    case "Домовладение":
                        {
                            dataGridView1.Rows[30].Cells[1].Value = lift.Text;
                            dataGridView1.Rows[43].Cells[1].Value = floor.Value.ToString();
                            dataGridView1.Rows[38].Cells[0].Value =
                           "Общая площадь домовладения, согласно правоустанавливающим документам (" +
                           registrationDoc.Text + "), в кв.м.";
                            dataGridView1.Rows[1].Cells[1].Value = fullAddressHouse();
                            dataGridView1.Rows[18].Cells[1].Value = houseType.Text;
                        }
                        break;

                    case "Земельный участок":
                        {
                            //objectDataGrid.Rows[1].Cells[1].Value = fullAddress();
                        }
                        break;

                    case "Домовладение с земельным участком":
                        {
                            dataGridView1.Rows[30].Cells[1].Value = lift.Text;
                            dataGridView1.Rows[43].Cells[1].Value = floor.Value.ToString();

                            dataGridView1.Rows[38].Cells[0].Value =
                            "Общая площадь домовладения, согласно правоустанавливающим документам (" +
                            registrationDoc.Text + "), в кв.м.";
                            dataGridView1.Rows[1].Cells[1].Value = fullAddressHouse();
                            dataGridView1.Rows[18].Cells[1].Value = houseType.Text;

                            if (!town.Text.Contains("г.Владикавказ"))
                            {
                                houseAnalogs.Rows[2].Cells[1].Value = town.Text;
                                houseAnalogs.Rows[2].Cells[2].Value = town.Text;
                                houseAnalogs.Rows[2].Cells[3].Value = town.Text;
                                houseAnalogs.Rows[2].Cells[4].Value = town.Text;

                                dirtGridAnalogs.Rows[0].Cells[1].Value = town.Text;
                                dirtGridAnalogs.Rows[0].Cells[2].Value = town.Text;
                                dirtGridAnalogs.Rows[0].Cells[3].Value = town.Text;
                                dirtGridAnalogs.Rows[0].Cells[4].Value = town.Text;
                            }
                            else
                            {
                                houseAnalogs.Rows[2].Cells[1].Value = street.Text;
                                houseAnalogs.Rows[2].Cells[2].Value = street.Text;
                                houseAnalogs.Rows[2].Cells[3].Value = street.Text;
                                houseAnalogs.Rows[2].Cells[4].Value = street.Text;

                                dirtGridAnalogs.Rows[0].Cells[1].Value = street.Text;
                                dirtGridAnalogs.Rows[0].Cells[2].Value = street.Text;
                                dirtGridAnalogs.Rows[0].Cells[3].Value = street.Text;
                                dirtGridAnalogs.Rows[0].Cells[4].Value = street.Text;
                            }
                        }
                        break;

                    default:
                        break;

                }
            }
            catch (Exception exp)
            {
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            calculateCost();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 2;
        }

        private void ownerSameCustomer_CheckedChanged(object sender, EventArgs e)
        {
            if (ownerSameCustomer.Checked)
            {
                ownerAddress.Text = customerAddres.Text;
                ownerInit.Text = customerInit.Text;
                ownerName.Text = customerName.Text;
                ownerPhone.Text = customerPhone.Text;
                ownerSurname.Text = customerSurname.Text;
                ownerPassport.Text = customerPassport.Text;
                ownerPassNum.Text = customerPassNum.Text;
                ownerPassOVD.Text = customerPassOVD.Text;
                ownerPassDate.Value = customerPassDate.Value;
                ownerAddress.Enabled = false;
                ownerInit.Enabled = false;
                ownerName.Enabled = false;
                ownerPhone.Enabled = false;
                ownerSurname.Enabled = false;
                ownerPassDate.Enabled = false;
                ownerPassNum.Enabled = false;
                ownerPassOVD.Enabled = false;
                ownerPassport.Enabled = false;
            }
            else
            {
                ownerAddress.Enabled = true;
                ownerInit.Enabled = true;
                ownerName.Enabled = true;
                ownerPhone.Enabled = true;
                ownerSurname.Enabled = true;
                ownerPassDate.Enabled = true;
                ownerPassNum.Enabled = true;
                ownerPassOVD.Enabled = true;
                ownerPassport.Enabled = true;
            }
        }

        /// <summary>
        ///     Return Header string
        /// </summary>
        public string topColontitulCreator()
        {
            string rooms = roomsAsString();
            string fullAddress;
            string buildNum = ".";

            if (buildingNum.Text != "")
            {
                buildNum = ", корп. " + buildingNum.Text + ".";
            }

            fullAddress = "Объект оценки - " + rooms + " квартира №" + appartmentNum.Text + " по адресу: " + town.Text +
                          ", " + street.Text + " " + houseNum.Text + buildNum;
            return fullAddress;
        }

        public string topColontitulCreatorHouse()
        {
            string rooms = roomsAsString();
            string fullAddress;
            string buildNum = ".";

            if (buildingNum.Text != "")
            {
                buildNum = ", корп. " + buildingNum.Text + ".";
            }

            fullAddress = "Частное домовладение и зем. участок по адресу: " + town.Text + ", " + street.Text + " " +
                          houseNum.Text + buildNum;
            return fullAddress;
        }

        /// <summary>
        ///     Convert roomsNum to String
        /// </summary>
        public string roomsAsString()
        {
            string rooms = null;
            switch (roomsNum.Value.ToString())
            {
                //ToDo проверить комнаты
                case "1":
                    rooms = "однокомнатная";
                    rooms1 = "однокомнатной";
                    roomsT = "Однокомнатная";
                    roomsN = "1 комнатная";
                    roomsX = "1-на комнатная квартира";
                    break;

                case "2":
                    rooms = "двухкомнатная";
                    rooms1 = "двухкомнатной";
                    roomsT = "Двухкомнатная";
                    roomsN = "2 комнатная";
                    roomsX = "2-ух комнатная квартира";
                    break;

                case "3":
                    rooms = "трехкомнатная";
                    rooms1 = "трехкомнатной";
                    roomsT = "Трехкомнатная";
                    roomsN = "3 комнатная";
                    roomsX = "3-ех комнатная квартира";
                    break;

                case "4":
                    rooms = "четырехкомнатная";
                    rooms1 = "четырехкомнатной";
                    roomsT = "Четырехкомнатная";
                    roomsN = "4 комнатная";
                    roomsX = "4-ех комнатная квартира";
                    break;

                case "5":
                    rooms = "пятикомнатная";
                    rooms1 = "пятикомнатной";
                    roomsT = "Пятикомнатная";
                    roomsN = "5 комнатная";
                    roomsX = "5-ти комнатная квартира";
                    break;

                case "6":
                    rooms = "шестикомнатная";
                    rooms1 = "шестикомнатной";
                    roomsT = "Шестикомнатная";
                    roomsN = "6 комнатная";
                    roomsX = "6-ти комнатная квартира";
                    break;

                //case "7": rooms = "семикомнатная"; break;
                //case '8': roomsNum = "однокомнатная": break;
                //case '9': roomsNum = "однокомнатная": break;
                default:
                    rooms = "";
                    break;
            }
            return rooms;
        }

        /// <summary>
        ///     Return Full Address String
        /// </summary>
        public string fullAddress()
        {
            string rooms = roomsAsString();
            string fullAddress;
            string buildNum = null;

            if (buildingNum.Text != "")
            {
                buildNum = ", корп. " + buildingNum.Text;
            }

            fullAddress = rooms + " квартира №" + appartmentNum.Text + " " + town.Text + ", " + street.Text + ", " +
                          houseNum.Text + buildNum;
            return fullAddress;
        }

        public string fullAddressHouse()
        {
            string rooms = roomsAsString();
            string fullAddress;
            string buildNum = null;

            if (buildingNum.Text != "")
            {
                buildNum = ", корп. " + buildingNum.Text;
            }

            fullAddress = " домовладение и земельный участок" + town.Text + ", " + street.Text + ", " + houseNum.Text + buildNum;
            return fullAddress;
        }

        public string fullAddressDirt()
        {
            string rooms = roomsAsString();
            string fullAddress;
            string buildNum = null;

            if (buildingNum.Text != "")
            {
                buildNum = ", корп. " + buildingNum.Text;
            }

            fullAddress = " земельный участок " + town.Text + ", " + street.Text + ", " + houseNum.Text + buildNum;
            return fullAddress;
        }

        public string getInits()
        {
            string name = "";
            string Init = "";

            if (customerName.Text != "")
            {
                name = customerName.Text.First() + ".";
            }

            if (customerInit.Text != "")
            {
                Init = customerInit.Text.First() + ".";
            }

            return (name + Init);
        }

        public void addObjectData()
        {
            objectDataGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            objectDataGrid.AutoResizeRows();
            objectDataGrid.AutoResizeColumns();

            objectDataGrid.Rows.Add("2.1.1", "Местоположение и окружение Объекта оценки"); //0
            objectDataGrid.Rows.Add("Местоположение Объекта оценки", " "); //1
            objectDataGrid.Rows.Add("Экологическая обстановка в районе", " "); //2
            objectDataGrid.Rows.Add("Интенсивность движения транспорта мимо дома", " "); //3
            objectDataGrid.Rows.Add("Транспортная доступность, обеспеченность общественным транспортом", " "); //4
            objectDataGrid.Rows.Add("Прилегающая транспортная магистраль, улица", " "); //5
            objectDataGrid.Rows.Add("Близость к скоростным магистралям, соседние улицы", " "); //6
            objectDataGrid.Rows.Add("Эстетичность окружающей застройки", " "); //7
            objectDataGrid.Rows.Add("Престижность района", " "); //8
            objectDataGrid.Rows.Add("Зонирование района (преобладающий тип застройки)", " "); //9
            objectDataGrid.Rows.Add("Близость к объектам социально-бытовой сферы", " "); //10
            objectDataGrid.Rows.Add("Близость к объектам развлечений и отдыха", " "); //11
            objectDataGrid.Rows.Add("Объекты промышленной инфраструктуры", " "); //12
            objectDataGrid.Rows.Add("Придомовая территория", " "); //13
            objectDataGrid.Rows.Add("Парковка возле дома", " "); //14
            objectDataGrid.Rows.Add("Наличие зеленых насаждений", " "); //15
            objectDataGrid.Rows.Add("Прочие особенности местоположения", " "); //16
            objectDataGrid.Rows.Add("2.1.2", "Описание дома, в котором расположена оцениваемая квартира"); //17
            objectDataGrid.Rows.Add("Тип дома", " "); //17
            objectDataGrid.Rows.Add("Год постройки", " "); //18
            objectDataGrid.Rows.Add("Этажность", " "); //19
            objectDataGrid.Rows.Add("Материал наружных стен", " "); //20
            objectDataGrid.Rows.Add("Материал перегородок", " "); //21
            objectDataGrid.Rows.Add("Группа капитальности", " "); //22
            objectDataGrid.Rows.Add("Наружная отделка", " "); //23
            objectDataGrid.Rows.Add("Состояние внеш.отделки, вид фасада", " "); //24
            objectDataGrid.Rows.Add("Характеристика перекрытий", " "); //25
            objectDataGrid.Rows.Add("Тип фундамента", " "); //26
            objectDataGrid.Rows.Add("Защищенность подъезда", " "); //27
            objectDataGrid.Rows.Add("Состояние обществ. зон подъезда", " "); //28
            objectDataGrid.Rows.Add("Лифт", " "); //29
            objectDataGrid.Rows.Add("Мусоропровод", " "); //30
            objectDataGrid.Rows.Add("Газ", " "); //31
            objectDataGrid.Rows.Add("Горячее водоснабжение", " "); //32
            objectDataGrid.Rows.Add("Отопление", " "); //33
            objectDataGrid.Rows.Add("Противопожарная безопасность", " "); //34
            objectDataGrid.Rows.Add("Наличие и тип парковки", " "); //35
            objectDataGrid.Rows.Add("Общее состояние дома", " "); //36
            objectDataGrid.Rows.Add("Наличие/ отсутствие дополнительных услуг для жильцов", " "); //37
            objectDataGrid.Rows.Add("Наличие/ отсутствие встроено-пристроенных нежилых помещений", " "); //38
            objectDataGrid.Rows.Add("2.1.3", "Описание оцениваемой квартиры"); //39
            objectDataGrid.Rows.Add(
                "Документ органа (организации), осуществившей технический учет и инвентаризацию Объекта оценки", " ");

            //40
            objectDataGrid.Rows.Add("Литер, согласно документа технического учета и инвентаризации", " "); //41
            objectDataGrid.Rows.Add("Этаж", " "); //42
            objectDataGrid.Rows.Add("Количество квартир на этаже", " "); //43
            objectDataGrid.Rows.Add("Тип планировки", " "); //44
            objectDataGrid.Rows.Add("Количество жил. комнат, их площадь", " ");
            objectDataGrid.Rows.Add(
                "Общая площадь (с учетом лоджий и балконов), согласно документа технического учета и инвентаризации, в кв.м.",
                " "); //45
            objectDataGrid.Rows.Add(
                "Общая площадь (без учета лоджий и балконов), согласно документа технического учета и инвентаризации, в кв.м.",
                " "); //46
            objectDataGrid.Rows.Add("Жилая площадь, согласно документа технич.учета и инвентаризации, в кв.м.", " ");

            //47
            objectDataGrid.Rows.Add("Площадь кухни, согласно документа технич.учета и инвентаризации, кв.м.", " "); //48
            objectDataGrid.Rows.Add("Санузел, количество санузлов", " "); //49
            objectDataGrid.Rows.Add("Балкон/лоджия, согласно документа технич.учета и инвентаризации", " "); //50
            objectDataGrid.Rows.Add(
                "Высота помещений по внутр. обмеру, согласно документа технического учета и инвентаризации, в м.", " ");

            //51
            objectDataGrid.Rows.Add(
                "Общая площадь квартиры, согласно правоустанавливающим документам (" + registrationDoc.Text +
                "), в кв.м.", " "); //52
            objectDataGrid.Rows.Add("Данные о неучтен. перепланировке", " "); //53
            objectDataGrid.Rows.Add("Остекление балкона/лоджии", " "); //54
            objectDataGrid.Rows.Add("Выход окон", " "); //55
            objectDataGrid.Rows.Add("Вспомогательные помещения", " "); //56
            objectDataGrid.Rows.Add("Смежные комнаты", " "); //57
            objectDataGrid.Rows.Add("Телефон", " "); //58
            objectDataGrid.Rows.Add("Дополн. системы безопасности", " "); //59
            objectDataGrid.Rows.Add("Система кондиционирования", " "); //60
            objectDataGrid.Rows.Add("Отделка: Полы", " "); //61
            objectDataGrid.Rows.Add("Отделка: Стены", " "); //62
            objectDataGrid.Rows.Add("Отделка: Потолки", " "); //63
            objectDataGrid.Rows.Add("Входная дверь", " "); //64
            objectDataGrid.Rows.Add("Межкомнатные двери", " "); //65
            objectDataGrid.Rows.Add("Окна", " "); //66
            objectDataGrid.Rows.Add("Сантехнические устройства", " "); //67
            objectDataGrid.Rows.Add("Подключение к электричеству", " "); //68
            objectDataGrid.Rows.Add("Подключение к холодному/горячему  водоснабжению", " "); //69
            objectDataGrid.Rows.Add("Подключение к канализации", " "); //70
            objectDataGrid.Rows.Add("Система отопления и отопительные приборы", " "); //71
            objectDataGrid.Rows.Add("Кухонная плита", " "); //72
            objectDataGrid.Rows.Add("Наличие следов протечек на потолке", " "); //73
            objectDataGrid.Rows.Add("Дополнительные удобства", " "); //74
            objectDataGrid.Rows.Add("Состояние отделки", " "); //75
            objectDataGrid.Rows.Add("Необходимые ремонтные работы", " "); //76
            objectDataGrid.Rows.Add("Текущее использование Объекта оценки", " "); //77
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (docTypeT == "Квартира")
            {
                var excelApp = new Microsoft.Office.Interop.Excel.Application();

                excelApp.Workbooks.Open(System.Windows.Forms.Application.StartupPath + "\\Черновик.xls", Missing,
                                        Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
                                        Missing, Missing, Missing, Missing);

                int i = 0;
                int j = 0;

                for (i = 0; i <= calculationAppartaments.RowCount - 1; i++)
                {
                    for (j = 2; j <= calculationAppartaments.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = calculationAppartaments[j, i];
                        excelApp.Workbooks[1].Sheets[1].Cells[i + 2, j + 1] = cell.Value;
                    }
                }

                excelApp.ActiveWorkbook.Save();
                excelApp.ActiveWorkbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\" + contractNum.Text + "Calc.xls");
                excelApp.ActiveWorkbook.Close(Missing);
                excelApp.Quit();

                excelApp = new Microsoft.Office.Interop.Excel.Application();

                excelApp.Workbooks.Open(System.Windows.Forms.Application.StartupPath + "\\analogs.xls", Missing, Missing,
                                        Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
                                        Missing, Missing, Missing);

                i = 0;
                j = 0;

                for (i = 0; i <= analogsGrid.RowCount - 1; i++)
                {
                    for (j = 1; j <= analogsGrid.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = analogsGrid[j, i];
                        excelApp.Workbooks[1].Sheets[1].Cells[i + 2, j + 1] = cell.Value;
                    }
                }

                excelApp.ActiveWorkbook.Save();
                excelApp.ActiveWorkbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\" + contractNum.Text + "Analogs.xls");
                excelApp.ActiveWorkbook.Close(Missing);
                excelApp.Quit();

                saveState();
            }
            System.Windows.Forms.Application.Exit();
        }

        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            objectDataGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            objectDataGrid.AutoResizeRows();
            objectDataGrid.AutoResizeColumns();
        }

        private string getUvaj()
        {
            var padeg = new Declension();
            int sex = padeg.GetSex(customerInit.Text);
            string cSex;
            if (sex == 1)
            {
                cSex = "м";
                uvaj = "Уважаемый";
            }
            else
            {
                cSex = "ж";
                uvaj = "Уважаемая";
            }

            return cSex;
        }

        private void customerPadeg()
        {
            var padeg = new Declension();

            int sex = padeg.GetSex(customerInit.Text);
            string cSex = getUvaj();
            customerFullNameR = padeg.GetFIOPadeg(customerSurname.Text, customerName.Text, customerInit.Text, cSex, 2);
            padeg.SeparateFIO(customerFullNameD, ref customerFamiliyR, ref customerNameR, ref customerSurnameR);
            customerFullNameD = padeg.GetFIOPadeg(customerSurname.Text, customerName.Text, customerInit.Text, cSex, 3);
            customerFullNameV = padeg.GetFIOPadeg(customerSurname.Text, customerName.Text, customerInit.Text, cSex, 4);
            customerFullNameT = padeg.GetFIOPadeg(customerSurname.Text, customerName.Text, customerInit.Text, cSex, 5);
            customerFullNameP = padeg.GetFIOPadeg(customerSurname.Text, customerName.Text, customerInit.Text, cSex, 6);
        }

        private void customerPadBut_Click(object sender, EventArgs e)
        {
            customerPadeg();
            new HouseCostCalculation.Padeg(customerFullNameR, customerFullNameD, customerFullNameV, customerFullNameT,
                                           customerFullNameP, 0);
        }

        /// <summary>
        ///     sets FullNameR
        /// </summary>
        public void fullNameRSet(string fullName, int t)
        {
            if (t == 1)
                ownerFullNameR = fullName;
            else
                customerFullNameR = fullName;
        }

        /// <summary>
        ///     sets FullNameD
        /// </summary>
        public void fullNameDSet(string fullName, int t)
        {
            if (t == 1)
                ownerFullNameD = fullName;
            else
                customerFullNameD = fullName;
        }

        private void ownerPadeg()
        {
            var padeg = new Declension();
            int sex = padeg.GetSex(ownerInit.Text);
            string cSex;
            if (sex == 1)
            {
                cSex = "м";
            }
            else
            {
                cSex = "ж";
            }

            ownerFullNameR = padeg.GetFIOPadeg(ownerSurname.Text, ownerName.Text, ownerInit.Text, cSex, 2);
            ownerFullNameD = padeg.GetFIOPadeg(ownerSurname.Text, ownerName.Text, ownerInit.Text, cSex, 3);
            ownerFullNameV = padeg.GetFIOPadeg(ownerSurname.Text, ownerName.Text, ownerInit.Text, cSex, 4);
            ownerFullNameT = padeg.GetFIOPadeg(ownerSurname.Text, ownerName.Text, ownerInit.Text, cSex, 5);
            ownerFullNameP = padeg.GetFIOPadeg(ownerSurname.Text, ownerName.Text, ownerInit.Text, cSex, 6);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            ownerPadeg();

            new HouseCostCalculation.Padeg(ownerFullNameR, ownerFullNameD, ownerFullNameV, ownerFullNameT,
                                           ownerFullNameP, 1);
        }

        public void addHouseData()
        {
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dataGridView1.AutoResizeRows();
            dataGridView1.AutoResizeColumns();
            dataGridView1.Rows.Add("2.1.1", "Местоположение и окружение Объекта оценки");
            dataGridView1.Rows.Add("Местоположение Объекта оценки", "");
            dataGridView1.Rows.Add("Экологическая обстановка в районе", "Хорошая");
            dataGridView1.Rows.Add("Интенсивность движения транспорта мимо дома", "Низкая");
            dataGridView1.Rows.Add("Транспортная доступность, обеспеченность общественным транспортом", "");
            dataGridView1.Rows.Add("Прилегающая транспортная магистраль, улица", "");
            dataGridView1.Rows.Add("Близость к скоростным магистралям, соседние улицы", "");
            dataGridView1.Rows.Add("Эстетичность окружающей застройки", "Удовлетворительно");
            dataGridView1.Rows.Add("Престижность района", "Удовлетворительно ");
            dataGridView1.Rows.Add("Зонирование района (преобладающий тип застройки)", "Жилая");
            dataGridView1.Rows.Add("Близость к объектам социально-бытовой сферы", "Неудовлетворительно");
            dataGridView1.Rows.Add("Близость к объектам развлечений и отдыха", "Неудовлетворительно");

            dataGridView1.Rows.Add("Придомовая территория", "Неохраняемая, не огороженная");

            dataGridView1.Rows.Add("Наличие зеленых насаждений", "Имеются");
            dataGridView1.Rows.Add("Прочие особенности местоположения", "Нет");
            dataGridView1.Rows.Add("Наличие расположенных рядом объектов, снижающих либо повышающих привлекательность",
                                   "Нет");

            dataGridView1.Rows.Add("2.1.2", "Описание оцениваемого домовладения");
            dataGridView1.Rows.Add("Год постройки", "Не установлен");
            dataGridView1.Rows.Add("Состояние", "Хорошее");
            dataGridView1.Rows.Add("Этажность", "Одноэтажный");
            dataGridView1.Rows.Add("Материал стен", "Кирпичные ");
            dataGridView1.Rows.Add("Материал перегородок", "Кирпичные ");
            dataGridView1.Rows.Add("Перекрытия", "Деревянные оштукатуренные ");
            dataGridView1.Rows.Add("Кровля", "Двускатная, шифер по деревянной обрешетке");
            dataGridView1.Rows.Add("Состояние внеш.отделки", "Удовлетворительное ");
            dataGridView1.Rows.Add("Внешний вид фасада дома", "Удовлетворительно ");
            dataGridView1.Rows.Add("Газоснабжение", "Есть");
            dataGridView1.Rows.Add("Горячее водоснабжение", "Автономное от газового котла (колонки) ");
            dataGridView1.Rows.Add("Отопление", "Автономное от газового котла ");
            dataGridView1.Rows.Add("Общее состояние дома", "Хорошее ");
            dataGridView1.Rows.Add(
                "Документ органа (организации), осуществившей технический учет и инвентаризацию Объекта оценки",
                "Кадастровый паспорт домовладения ГУП «Аланиятехинвентаризации РСО-Алания» по инв.№697 от 07/07/08г");
            dataGridView1.Rows.Add("Литер(а), согласно документа технического учета и инвентаризации", "Литер «А»");
            dataGridView1.Rows.Add("Тип планировки", "Фиксированный");
            dataGridView1.Rows.Add("Количество жилых комнат, площадь ",
                                   "Пять  жилых комнат: 24,8м2; 12,9м2; 21,9м2; 15,4м2 и 12,0");
            dataGridView1.Rows.Add(
                "Общая площадь (с учетом лоджий и балконов), согласно документа технического учета и инвентаризации, в кв.м.",
                "144,7");
            dataGridView1.Rows.Add(
                "Общая площадь (без учета лоджий и балконов), согласно документа технического учета и инвентаризации, в кв.м.",
                "144,7");
            dataGridView1.Rows.Add("Жилая площадь, согласно документа технич.учета и инвентаризации, в кв.м.", "87,0");
            dataGridView1.Rows.Add(
                "Высота помещений по внутр. обмеру, согласно документа технического учета и инвентаризации, в м.",
                "2,9м, 2,4");
            dataGridView1.Rows.Add(
                "Общая площадь, согласно правоустанавливающим документам (Свидетельство о государственной регистрации права Управления Федеральной регистрационной службы по РСО-Алания серия 15 АЕ №706443 от 24/08/05г.), в кв.м.",
                "144,7");
            dataGridView1.Rows.Add("Санузел, количество санузлов", "Один совмещенный, общей площадью 9,9");
            dataGridView1.Rows.Add("Смежные комнаты", "Нет  ");
            dataGridView1.Rows.Add("Телефон", "Есть   ");
            dataGridView1.Rows.Add("Отделка: Полы",
                                   "В жилых комнатах деревянные, на кухне и в сан. узле плиточные. Состояние хорошее");
            dataGridView1.Rows.Add("Отделка: Стены",
                                   "В жилых комнатах оштукатурено, побелено, в рабочей части кухни и в сан. узле плиточные. Состояние хорошее ");
            dataGridView1.Rows.Add("Отделка: Потолки", "Оштукатурено, побелено. Состояние хорошее ");
            dataGridView1.Rows.Add("Межкомнатные двери", "Деревянные полотна и филенчатые. Состояние хорошее ");
            dataGridView1.Rows.Add("Окна", "Деревянные рамы, двустворчатые, двойное остекление. Состояние хорошее");
            dataGridView1.Rows.Add("Сантехника", "Полностью установлены ");
            dataGridView1.Rows.Add("Подключение к электричеству", "Есть ");
            dataGridView1.Rows.Add("Подключение к холодному/горячему  водоснабжению",
                                   "Холодное водоснабжение от сельских сетей, горячее водоснабжение автономное от газовой колонки. Трубы и запорная арматура металлич., состояние удовлетворит.");
            dataGridView1.Rows.Add("Подключение к канализации", "Канализация ");
            dataGridView1.Rows.Add("Отопительные приборы",
                                   "Простые металлические радиаторы отопления. Состояние хорошее  ");
            dataGridView1.Rows.Add("Кухонная плита", "Отечественная, газовая четырехкомфорочная ");
            dataGridView1.Rows.Add("Наличие следов протечек на потолке", "Нет ");
            dataGridView1.Rows.Add("Наличие перепланировки", "Не выявлено");
            dataGridView1.Rows.Add("Дополнительные удобства", "Нет ");
            dataGridView1.Rows.Add("Состояние отделки", "Хорошее  ");
            dataGridView1.Rows.Add("Необходимые ремонтные работы", "Необходимо проведение косметических работ ");
            dataGridView1.Rows.Add("Текущее использование Объекта оценки",
                                   "Некоммерческое использование, жилое домовладение, проживание. ");
        }

        public void addGridData()
        {
            dirtCalcGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dirtCalcGrid.AutoResizeRows();
            dirtCalcGrid.AutoResizeColumns();
            dirtCalcGrid.Rows.Add("Адрес объекта", "г.Владикавказ, «Иристон»", "г.Владикавказ, «Иристон»",
                                  "г.Владикавказ, «Иристон»"); //1
            dirtCalcGrid.Rows.Add("Цена предложения за участок, руб.", "350 000", "350 000", "350 000"); //2
            dirtCalcGrid.Rows.Add("Площадь участка, сот.", "6", "6", "6"); //3
            dirtCalcGrid.Rows.Add("Цена предложения за 1 сот., руб./сот.", "", "", ""); //4
            dirtCalcGrid.Rows.Add("Перевод предложения в цену сделки (поправка на торг)", "0,95", "0,95", "0,95"); //5
            dirtCalcGrid.Rows.Add("Вид права собственности", "Полное право", "Полное право", "Полное право"); //6
            dirtCalcGrid.Rows.Add("Поправка на право собственности", "1", "1", "1"); //7
            dirtCalcGrid.Rows.Add("Условия финансовых расчетов", "За собственные средства в момент оформления",
                                  "За собственные средства в момент оформления",
                                  "За собственные средства в момент оформления"); //8
            dirtCalcGrid.Rows.Add("Поправка на условия финансовых расчетов", "1", "1", "1"); //9
            dirtCalcGrid.Rows.Add("Условия продажи", "Свободная продажа", "Свободная продажа", "Свободная продажа");

            //10
            dirtCalcGrid.Rows.Add("Поправка на условия продажи", "1", "1", "1"); //11
            dirtCalcGrid.Rows.Add("Дата предложения", "Август 2011г.", "Август 2011г.", "Август 2011г."); //12
            dirtCalcGrid.Rows.Add("Поправка на дату предложения", "1", "1", "1"); //13
            dirtCalcGrid.Rows.Add("Район расположение", "Окраина села", "Окраина села", "Окраина села"); //14
            dirtCalcGrid.Rows.Add("Поправка на район расположение", "1", "1", "1"); //15
            dirtCalcGrid.Rows.Add("Целевое назначение", "Для эксплуатации жилого дома", "Для эксплуатации жилого дома",
                                  "Для эксплуатации жилого дома"); //16
            dirtCalcGrid.Rows.Add("Поправка на целевое назначение", "1", "1", "1"); //17
            dirtCalcGrid.Rows.Add("Размер участка (масштаб участка), в сот.", "6,00", "6,00", "6,00"); //18
            dirtCalcGrid.Rows.Add("Поправка на размер участка (масштаб участка)", "1", "1", "1"); //19
            dirtCalcGrid.Rows.Add("Наличие коммуникаций", "Все коммуникации", "Все коммуникации", "Все коммуникации");

            //20
            dirtCalcGrid.Rows.Add("Поправка на наличие коммуникаций", "1", "1", "1"); //21
            dirtCalcGrid.Rows.Add("Наличие и состояние подъездных путей (дороги)", "Хорошо", "Хорошо", "Хорошо"); //22
            dirtCalcGrid.Rows.Add("Поправка на наличие и состояние подъездных путей", "1", "1", "1"); //23
            dirtCalcGrid.Rows.Add("Рельеф и форма участка", "Рельеф ровный, форма прямоуг.",
                                  "Рельеф ровный, форма прямоуг.", "Рельеф ровный, форма прямоуг."); //24
            dirtCalcGrid.Rows.Add("Поправка на рельеф и форму участка", "1", "1", "1"); //25
            dirtCalcGrid.Rows.Add("Итоговая скорректированная стоимость аналога, руб./сот.", "1", "1", "1"); //26
            dirtCalcGrid.Rows.Add("Количество произведенных корректировок, коррект.", "1", "1", "1"); //27
            dirtCalcGrid.Rows.Add("Весовой коэффициент, в зависимости от кол-ва произв. корректировок, доля един. ", "1",
                                  "1", "1"); //28
            dirtCalcGrid.Rows.Add("Скорректированная стоимость, доля в итоговой  стоимости, руб./сот.", "1", "1", "1");

            //29
            dirtCalcGrid.Rows.Add("Итоговая стоимость 1 сотки оценив. зем. участка, руб./сот.", "1", "", ""); //30
            dirtCalcGrid.Rows.Add("Общая площадь оцениваемого земельного участка, сот.", "1", "", ""); //31
            dirtCalcGrid.Rows.Add(" Итоговая стоимость оцениваемого зем. участка, руб.", "1", "", ""); //32
            dirtCalcGrid.Rows.Add(" Итоговая стоимость оцениваемого зем. участка с учетом округления, тыс. руб.", "1",
                                  "", ""); //33
            dirtCalcGrid.Rows.Add("Ликвидационная стоимость оцениваемого зем. участка с учетом округления, тыс. руб.",
                                  "1", "", ""); //34
        }

        private DataTable getDataFromXLS(string strFilePath)
        {
            try
            {
                string strConnectionString = "";
                strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=" + strFilePath + "; Jet OLEDB:Engine Type=5;" +
                                      "Extended Properties=Excel 8.0;";
                var cnCSV = new OleDbConnection(strConnectionString);
                cnCSV.Open();
                var cmdSelect = new OleDbCommand(@"SELECT * FROM [Лист1$]", cnCSV);
                var daCSV = new OleDbDataAdapter();
                daCSV.SelectCommand = cmdSelect;
                var dtCSV = new DataTable();
                daCSV.Fill(dtCSV);
                cnCSV.Close();
                daCSV = null;
                return dtCSV;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
            }
        }

        public void calculateCost()
        {
            try
            {
                string cellValue;

                //todo: Добавить цикл для обхода всех столбцов
                cost_count1 = 0;
                cost_count2 = 0;
                cost_count3 = 0;
                int analogsCount = calculationAppartaments.ColumnCount - 3;
                cost_count1 = CalcCost(2);
                cost_count2 = CalcCost(3);
                cost_count3 = CalcCost(4);

                //Аналог 2

                //Final costs
                double koef_count, t2;
                koef_count = cost_count1 + cost_count2 + cost_count3;
                t2 = 1 / koef_count;
                cellValue = calculationAppartaments.Rows[33].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cost_cor_koef1 = double.Parse(cellValue);
                }
                cellValue = calculationAppartaments.Rows[33].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cost_cor_koef2 = double.Parse(cellValue);
                }
                cellValue = calculationAppartaments.Rows[33].Cells[4].Value.ToString();
                if (cellValue != "")
                {
                    cost_cor_koef3 = double.Parse(cellValue);
                }

                cellValue = calculationAppartaments.Rows[31].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost113 = double.Parse(cellValue);
                }
                cellValue = calculationAppartaments.Rows[31].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost213 = double.Parse(cellValue);
                }
                cellValue = calculationAppartaments.Rows[31].Cells[4].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost313 = double.Parse(cellValue);
                }

                cor_cost_final1 = Math.Round(cor_cost113 * cost_cor_koef1);
                calculationAppartaments.Rows[34].Cells[2].Value = cor_cost_final1.ToString();

                if (cost_cor_koef2 != 0)
                {
                    cor_cost_final2 = Math.Round(cor_cost213 * cost_cor_koef2);
                }
                else
                {
                    cor_cost_final2 = 0;
                }

                calculationAppartaments.Rows[34].Cells[3].Value = cor_cost_final2.ToString();
                if (cost_cor_koef3 != 0)
                {
                    cor_cost_final3 = Math.Round(cor_cost313 * cost_cor_koef3);
                }
                else
                {
                    cor_cost_final3 = 0;
                }

                calculationAppartaments.Rows[34].Cells[4].Value = cor_cost_final3.ToString();
                final_cost_m = Math.Round(cor_cost_final1 + cor_cost_final2 + cor_cost_final3);
                calculationAppartaments.Rows[35].Cells[2].Value = final_cost_m.ToString();

                finalCost = Math.Round(final_cost_m * m_final);
                calculationAppartaments.Rows[37].Cells[2].Value = finalCost.ToString();

                finalCostRounded = Math.Round(finalCost / 1000);
                calculationAppartaments.Rows[38].Cells[2].Value = finalCostRounded.ToString();
                costStr = RusCurrency.Str(finalCostRounded * 1000, "RUR");
                costStr = costStr.Replace("00 копеек", "");
                likvidCost = Math.Round(finalCostRounded * 0.66);
                calculationAppartaments.Rows[39].Cells[2].Value = likvidCost.ToString();

                //date1 = contractDate.Text;
                //if (date1 != "")
                //{
                //    dataGridView2.Rows[5].Cells[2].ValueType = System.Type.;
                //    dataGridView2.Rows[5].Cells[2].Value = date1.ToString();
                //}
            }
            catch (Exception exp)
            {
            }
        }

        private int CalcCost(int i)
        {
            try
            {
                cost_count1 = 0;
                string cellValue;
                cellValue = calculationAppartaments.Rows[0].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cost1 = double.Parse(cellValue);
                }
                cellValue = calculationAppartaments.Rows[1].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    m1 = double.Parse(cellValue);
                }
                cellValue = calculationAppartaments.Rows[3].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_torg = double.Parse(cellValue);
                    if (cor_torg != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[6].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor1 = double.Parse(cellValue);
                    if (cor1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[8].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_place1 = double.Parse(cellValue);
                    if (cor_place1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[10].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_type1 = double.Parse(cellValue);
                    if (cor_type1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[12].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_date1 = double.Parse(cellValue);
                    if (cor_date1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[14].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_floor1 = double.Parse(cellValue);
                    if (cor_floor1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[16].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_m1 = double.Parse(cellValue);
                    if (cor_m1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[18].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_b1 = double.Parse(cellValue);
                    if (cor_b1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[20].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_height1 = double.Parse(cellValue);
                    if (cor_height1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[22].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_class1 = double.Parse(cellValue);
                    if (cor_class1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[24].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_phone1 = double.Parse(cellValue);
                    if (cor_phone1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[26].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_com1 = double.Parse(cellValue);
                    if (cor_com1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[28].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_t1 = double.Parse(cellValue);
                    if (cor_t1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[30].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_lift1 = double.Parse(cellValue);
                    if (cor_lift1 != 1.00)
                    {
                        cost_count1++;
                    }
                }
                cellValue = calculationAppartaments.Rows[36].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    m_final = double.Parse(cellValue);
                }

                cost_m1 = Math.Round(cost1 / m1);
                calculationAppartaments.Rows[2].Cells[i].Value = cost_m1.ToString();
                cor_cost1 = Math.Round(cost_m1 * cor_torg);
                calculationAppartaments.Rows[4].Cells[i].Value = cor_cost1.ToString();

                cor_cost11 = Math.Round(cor1 * cor_cost1);
                calculationAppartaments.Rows[7].Cells[i].Value = cor_cost11.ToString();

                cor_cost12 = Math.Round(cor_cost11 * cor_place1);
                calculationAppartaments.Rows[9].Cells[i].Value = cor_cost12.ToString();

                cor_cost13 = Math.Round(cor_cost12 * cor_type1);
                calculationAppartaments.Rows[11].Cells[i].Value = cor_cost13.ToString();

                cor_cost14 = Math.Round(cor_cost13 * cor_date1);
                calculationAppartaments.Rows[13].Cells[i].Value = cor_cost14.ToString();

                cor_cost15 = Math.Round(cor_cost14 * cor_floor1);
                calculationAppartaments.Rows[15].Cells[i].Value = cor_cost15.ToString();

                cor_cost16 = Math.Round(cor_cost15 * cor_m1);
                calculationAppartaments.Rows[17].Cells[i].Value = cor_cost16.ToString();

                cor_cost17 = Math.Round(cor_cost16 * cor_b1);
                calculationAppartaments.Rows[19].Cells[i].Value = cor_cost17.ToString();

                cor_cost18 = Math.Round(cor_cost17 * cor_height1);
                calculationAppartaments.Rows[21].Cells[i].Value = cor_cost18.ToString();

                cor_cost19 = Math.Round(cor_cost18 * cor_class1);
                calculationAppartaments.Rows[23].Cells[i].Value = cor_cost19.ToString();

                cor_cost110 = Math.Round(cor_cost19 * cor_phone1);
                calculationAppartaments.Rows[25].Cells[i].Value = cor_cost110.ToString();

                cor_cost111 = Math.Round(cor_cost110 * cor_com1);
                calculationAppartaments.Rows[27].Cells[i].Value = cor_cost111.ToString();

                cor_cost112 = Math.Round(cor_cost111 * cor_t1);
                calculationAppartaments.Rows[29].Cells[i].Value = cor_cost112.ToString();

                cor_cost113 = Math.Round(cor_cost112 * cor_lift1);
                calculationAppartaments.Rows[31].Cells[i].Value = cor_cost113.ToString();

                calculationAppartaments.Rows[32].Cells[i].Value = cost_count1.ToString();
                return cost_count1;
            }
            catch (Exception exp)
            {
                return 0;
            }
        }

        public void calculateCostHouse()
        {
            try
            {
                string cellValue;
                cost_count1 = calcHouseAnalog(2);
                cost_count2 = calcHouseAnalog(3);
                cost_count3 = calcHouseAnalog(4);
                cellValue = houseCalcGrid.Rows[37].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    m_final = double.Parse(cellValue);
                }
                cellValue = houseCalcGrid.Rows[32].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost113 = double.Parse(cellValue);
                }
                cellValue = houseCalcGrid.Rows[32].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost213 = double.Parse(cellValue);
                }
                cellValue = houseCalcGrid.Rows[32].Cells[4].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost313 = double.Parse(cellValue);
                }

                //Final costs
                cellValue = houseCalcGrid.Rows[34].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cost_cor_koef1 = double.Parse(cellValue);
                }
                cellValue = houseCalcGrid.Rows[34].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cost_cor_koef2 = double.Parse(cellValue);
                }
                cellValue = houseCalcGrid.Rows[34].Cells[4].Value.ToString();
                if (cellValue != "")
                {
                    cost_cor_koef3 = double.Parse(cellValue);
                }

                double koef_count, t2;
                koef_count = cost_count1 + cost_count2 + cost_count3;
                t2 = 1 / koef_count;

                cor_cost_final1 = Math.Round(cor_cost113 * cost_cor_koef1);
                houseCalcGrid.Rows[35].Cells[2].Value = cor_cost_final1.ToString();

                cor_cost_final2 = Math.Round(cor_cost213 * cost_cor_koef2);
                houseCalcGrid.Rows[35].Cells[3].Value = cor_cost_final2.ToString();

                cor_cost_final3 = Math.Round(cor_cost313 * cost_cor_koef3);
                houseCalcGrid.Rows[35].Cells[4].Value = cor_cost_final3.ToString();

                final_cost_m = Math.Round(cor_cost_final1 + cor_cost_final2 + cor_cost_final3);
                houseCalcGrid.Rows[36].Cells[2].Value = final_cost_m.ToString();

                finalCost = Math.Round(final_cost_m * m_final);
                houseCalcGrid.Rows[38].Cells[2].Value = finalCost.ToString();

                finalCostRounded = Math.Round(finalCost / 1000);
                houseCalcGrid.Rows[39].Cells[2].Value = finalCostRounded.ToString();
                costStr = RusCurrency.Str((finalCostRounded + finalDirtCost / 1000) * 1000);
                costStr = costStr.Replace("00 копеек", "");
                likvidCost = Math.Round(finalCostRounded * 0.66);
                houseCalcGrid.Rows[40].Cells[2].Value = likvidCost.ToString();

                //date1 = contractDate.Text;
                //if (date1 != "")
                //{
                //    dataGridView5.Rows[5].Cells[2].ValueType = System.Type.;
                //    dataGridView5.Rows[5].Cells[2].Value = date1.ToString();
                //}
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message);
            }
        }

        private int calcHouseAnalog(int i)
        {
            try
            {
                string cellValue;
                int costCount = 0;
                cellValue = houseCalcGrid.Rows[0].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cost1 = double.Parse(cellValue);
                }
                cellValue = houseCalcGrid.Rows[1].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    m1 = double.Parse(cellValue);
                }
                cellValue = houseCalcGrid.Rows[3].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_torg = double.Parse(cellValue);
                    if (cor_torg != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[6].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor1 = double.Parse(cellValue);
                    if (cor1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[9].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_place1 = double.Parse(cellValue);
                    if (cor_place1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[11].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_type1 = double.Parse(cellValue);
                    if (cor_type1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[13].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_date1 = double.Parse(cellValue);
                    if (cor_date1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[15].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_floor1 = double.Parse(cellValue);
                    if (cor_floor1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[17].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_m1 = double.Parse(cellValue);
                    if (cor_m1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[19].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_b1 = double.Parse(cellValue);
                    if (cor_b1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[21].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_height1 = double.Parse(cellValue);
                    if (cor_height1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[23].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_class1 = double.Parse(cellValue);
                    if (cor_class1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[25].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_phone1 = double.Parse(cellValue);
                    if (cor_phone1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[27].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_com1 = double.Parse(cellValue);
                    if (cor_com1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[29].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_t1 = double.Parse(cellValue);
                    if (cor_t1 != 1.00)
                    {
                        costCount++;
                    }
                }
                cellValue = houseCalcGrid.Rows[31].Cells[i].Value.ToString();
                if (cellValue != "")
                {
                    cor_lift1 = double.Parse(cellValue);
                    if (cor_lift1 != 1.00)
                    {
                        costCount++;
                    }
                }

                //cost_m1 = Math.Round(cost1 / m1);
                //houseCalcGrid.Rows[2].Cells[i].Value = cost_m1.ToString();
                cor_cost1 =
                    Math.Round(double.Parse(houseCalcGrid.Rows[0].Cells[i].Value.ToString()) -
                               double.Parse(houseCalcGrid.Rows[2].Cells[i].Value.ToString()));
                houseCalcGrid.Rows[3].Cells[i].Value = cor_cost1.ToString();

                cost_m1 = Math.Round(cor_cost1 / double.Parse(houseCalcGrid.Rows[4].Cells[i].Value.ToString()));
                houseCalcGrid.Rows[5].Cells[i].Value = cost_m1;

                cor_cost11 = Math.Round(cor1 * cost_m1);
                houseCalcGrid.Rows[7].Cells[i].Value = cor_cost11.ToString();

                cor_cost12 = Math.Round(cor_cost11 * cor_place1);
                houseCalcGrid.Rows[10].Cells[i].Value = cor_cost12.ToString();

                cor_cost13 = Math.Round(cor_cost12 * cor_type1);
                houseCalcGrid.Rows[12].Cells[i].Value = cor_cost13.ToString();

                cor_cost14 = Math.Round(cor_cost13 * cor_date1);
                houseCalcGrid.Rows[14].Cells[i].Value = cor_cost14.ToString();

                cor_cost15 = Math.Round(cor_cost14 * cor_floor1);
                houseCalcGrid.Rows[16].Cells[i].Value = cor_cost15.ToString();

                cor_cost16 = Math.Round(cor_cost15 * cor_m1);
                houseCalcGrid.Rows[18].Cells[i].Value = cor_cost16.ToString();

                cor_cost17 = Math.Round(cor_cost16 * cor_b1);
                houseCalcGrid.Rows[20].Cells[i].Value = cor_cost17.ToString();

                cor_cost18 = Math.Round(cor_cost17 * cor_height1);
                houseCalcGrid.Rows[22].Cells[i].Value = cor_cost18.ToString();

                cor_cost19 = Math.Round(cor_cost18 * cor_class1);
                houseCalcGrid.Rows[24].Cells[i].Value = cor_cost19.ToString();

                cor_cost110 = Math.Round(cor_cost19 * cor_phone1);
                houseCalcGrid.Rows[26].Cells[i].Value = cor_cost110.ToString();

                cor_cost111 = Math.Round(cor_cost110 * cor_com1);
                houseCalcGrid.Rows[28].Cells[i].Value = cor_cost111.ToString();

                cor_cost112 = Math.Round(cor_cost111 * cor_t1);
                houseCalcGrid.Rows[30].Cells[i].Value = cor_cost112.ToString();

                cor_cost113 = Math.Round(cor_cost112 * cor_lift1);
                houseCalcGrid.Rows[32].Cells[i].Value = cor_cost113.ToString();

                houseCalcGrid.Rows[33].Cells[i].Value = costCount.ToString();
                return costCount;
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message);
                return 0;
            }
        }

        private void calculationAppartaments_CellStateChanged(object sender, DataGridViewCellStateChangedEventArgs e)
        {
        }

        private void floors_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                updateTables();
            }
            catch (Exception exp)
            {
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            addObjectData();
        }

        private void SaveHouse(object sender, EventArgs e)
        {
            try
            {
                //HouseCostCalculation.House h = new HouseCostCalculation.House();
                //h.saveHouse(this);
                string townName = " " + town.Text + ", ";

                if ((town.Text == "г. Владикавказ") || (town.Text == "г.Владикавказ"))
                {
                    townName = " ";
                }

                string buildNum = null;

                if (buildingNum.Text != "")
                {
                    buildNum = "корп. " + buildingNum.Text;
                }

                roomsAsString();
                string fileName = "отчет  " + contractNum.Text + " от " + calculationDate.Text + " договор от" +
                                  contractDate.Text + " " + fullAddressHouse() + " " + ownerSurname.Text + " " +
                                  ownerName.Text + " для " + customerSurname.Text + " " + customerName.Text + " " +
                                  bankName.Text;
                saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();

                if (DialogResult.OK == saveFileDialog1.ShowDialog())
                {
                    wdApp = new Application();
                    var wdDoc = new Document();

                    wdDoc = wdApp.Documents.Open(System.Windows.Forms.Application.StartupPath + "\\шаблоны\\Дом.doc",
                                                 Missing, true);
                    object replaceAll = WdReplace.wdReplaceAll;

                    // Gets a NumberFormatInfo associated with the en-US culture.
                    NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;

                    nfi.NumberDecimalDigits = 0;
                    nfi.NumberGroupSeparator = " ";

                    nfi.PositiveSign = "";

                    string ownerFullName = ownerSurname.Text + " " + ownerName.Text + " " + ownerInit.Text;
                    string customerFullName = customerSurname.Text + " " + customerName.Text + " " + customerInit.Text;
                    double dirtCost = double.Parse(dirtCalcGrid.Rows[31].Cells[1].Value.ToString());
                    calculationDate.CustomFormat = "dd MMMM yyyy";
                    string calculationDateStr = calculationDate.Text;
                    int lenght = calculationDateStr.Length;
                    string temp = null;
                    string t;

                    for (int i = 0; i < lenght; i++)
                    {
                        if (i == 3)
                        {
                            t = calculationDateStr[i].ToString().ToUpper();
                            temp += t;
                        }
                        else
                        {
                            temp += calculationDateStr[i];
                        }
                    }

                    calculationDateStr = temp;

                    calculationDate.CustomFormat = "dd/MM/yy";
                    int sentencesCount = wdDoc.Sentences.Count;
                    string topColontitul = topColontitulCreatorHouse();

                    wdDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = topColontitul;

                    int count = houseAnalogs.RowCount - 1;

                    //Объект оценки

                    AddHouseAnalog(count, 0);

                    //Аналог 1
                    AddHouseAnalog(count, 1);

                    //Аналог 2
                    AddHouseAnalog(count, 2);

                    //Аналог 3
                    AddHouseAnalog(count, 3);

                    int dirtAnalogsCount = dirtCalcGrid.RowCount - 1;

                    //Аналог 1
                    AddGridCost(dirtAnalogsCount, 1);

                    //Аналог 2
                    AddGridCost(dirtAnalogsCount, 2);

                    //Аналог 3
                    AddGridCost(dirtAnalogsCount, 3);

                    ReplaceTextWord(ref wdApp, "@@MO@@", MO.Text);
                    ReplaceTextWord(ref wdApp, "@@dirtCost@@", dirtCalcGrid.Rows[31].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@dirtm2@@", dirtm2.Text);
                    ReplaceTextWord(ref wdApp, "@@dirtCostR@@", dirtCalcGrid.Rows[32].Cells[1].Value.ToString());

                    ReplaceTextWord(ref wdApp, "@@kadastrNum@@", dirtKadastr.Text);
                    ReplaceTextWord(ref wdApp, "@@dirtDoc@@", gridDoc.Text);

                    ReplaceTextWord(ref wdApp, "@@calculationDateStr@@", calculationDateStr);

                    ReplaceTextWord(ref wdApp, "@@houseType@@", houseType.Text.ToLower());

                    ReplaceTextWord(ref wdApp, "@@roomsT@@", roomsT);

                    ReplaceTextWord(ref wdApp, "@@roomsX@@", roomsX);

                    ReplaceTextWord(ref wdApp, "@@lm2@@", lm2text.Text);

                    ReplaceTextWord(ref wdApp, "@@m2@@", m2text.Text);
                    ReplaceTextWord(ref wdApp, "@@customerNameInits@@", customerFamiliyR + " " + getInits());
                    ReplaceTextWord(ref wdApp, "@@calculationDate@@", calculationDate.Text); ReplaceTextWord(ref wdApp, "@@ownerFullname@@", ownerFullName); ReplaceTextWord(ref wdApp, "@@customerFullname@@", customerFullName); roomsAsString(); ReplaceTextWord(ref wdApp, "@@rooms1@@", rooms1); ReplaceTextWord(ref wdApp, "@@ownerFullnameR@@", ownerFullNameR); ReplaceTextWord(ref wdApp, "@@customerFullnameR@@", customerFullNameR); ReplaceTextWord(ref wdApp, "@@customerFullnameT@@", customerFullNameT); ReplaceTextWord(ref wdApp, "@@ownerFullnameD@@", ownerFullNameD); ReplaceTextWord(ref wdApp, "@@ownerFullnameT@@", ownerFullNameT); ReplaceTextWord(ref wdApp, "@@customerFullnameD@@", customerFullNameD); ReplaceTextWord(ref wdApp, "@@rooms@@", roomsAsString()); ReplaceTextWord(ref wdApp, "@@appartmentNum@@", "№" + appartmentNum.Text); ReplaceTextWord(ref wdApp, "@@street@@", street.Text); ReplaceTextWord(ref wdApp, "@@houseNum@@", houseNum.Text);

                    buildNum = null;
                    if (buildingNum.Text != "")
                    {
                        buildNum = " корп." + buildingNum.Text;
                    }
                    else
                    {
                        buildNum = buildingNum.Text;
                    }

                    ReplaceTextWord(ref wdApp, "@@buildingNum@@",
                                    buildNum
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@customerAddress@@",
                                    customerAddres.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@floor@@",
                                    floor.Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@floors@@", floors.Text);

                    ReplaceTextWord(ref wdApp, "@@town@@",
                                    town.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@cost@@",
                                    finalCostRounded.ToString("N", nfi)
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@costFull@@",
                                    (finalCostRounded + dirtCost / 1000).ToString("N", nfi)
                        )
                        ;
                    ReplaceTextWord(ref wdApp, "@@likvidCostDirt@@", dirtCalcGrid.Rows[33].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@likvidCostFull@@", (likvidCostDirt + likvidCost).ToString());

                    ReplaceTextWord(ref wdApp, "@@contractNum@@",
                                    contractNum.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@contractDate@@", contractDate.Text);

                    ReplaceTextWord(ref wdApp, "@@customerName@@",
                                    customerName.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@customerInit@@",
                                    customerInit.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@likvidCost@@",
                                    likvidCost.ToString("N", nfi)
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@stringCost@@",
                                    costStr.ToLower()
                        )
                        ;

                    getUvaj();

                    ReplaceTextWord(ref wdApp, "@@uvaj@@",
                                    uvaj
                        )
                        ;

                    //Customer Passport

                    ReplaceTextWord(ref wdApp, "@@customerPassport@@",
                                    customerPassport.Text)
                        ;

                    ReplaceTextWord(ref wdApp, "@@customerPassNum@@",
                                    customerPassNum.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@customerPassOVD@@",
                                    customerPassOVD.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@customerPassDate@@",
                                    customerPassDate.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@customerFullAddress@@",
                                    customerAddres.Text
                        )
                        ;

                    //owner Passport

                    ReplaceTextWord(ref wdApp, "@@ownerPassport@@", ownerPassport.Text);
                    ReplaceTextWord(ref wdApp, "@@ownerPassNum@@",
                                    ownerPassNum.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@ownerPassOVD@@", ownerPassOVD.Text);

                    ReplaceTextWord(ref wdApp, "@@ownerPassDate@@",
                                    ownerPassDate.Text
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@ownerFullAddress@@",
                                    ownerAddress.Text)
                        ;

                    ReplaceTextWord(ref wdApp, "@@ownerDoc@@", ownerDocs.Text);

                    ReplaceTextWord(ref wdApp, "@@registrationDoc@@", registrationDoc.Text);

                    ReplaceTextWord(ref wdApp, "@@tehPass@@", dataGridView1.Rows[41].Cells[1].Value.ToString());

                    ReplaceTextWord(ref wdApp, "@@2.1.1.2@@",
                                    dataGridView1.Rows[2].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.3@@",
                                    dataGridView1.Rows[3].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.4@@",
                                    dataGridView1.Rows[4].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.5@@",
                                    dataGridView1.Rows[5].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.6@@",
                                    dataGridView1.Rows[6].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.7@@",
                                    dataGridView1.Rows[7].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.8@@",
                                    dataGridView1.Rows[8].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.9@@",
                                    dataGridView1.Rows[9].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.10@@",
                                    dataGridView1.Rows[10].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.11@@",
                                    dataGridView1.Rows[11].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.12@@",
                                    dataGridView1.Rows[12].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.13@@",
                                    dataGridView1.Rows[13].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.14@@",
                                    dataGridView1.Rows[14].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.1.15@@",
                                    dataGridView1.Rows[15].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.1@@",
                                    dataGridView1.Rows[17].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.2@@",
                                    dataGridView1.Rows[18].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.3@@"
                                    ,
                                    dataGridView1.Rows[19].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.4@@"
                                    ,
                                    dataGridView1.Rows[20].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.5@@"
                                    ,
                                    dataGridView1.Rows[21].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.6@@"
                                    ,
                                    dataGridView1.Rows[22].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.7@@"
                                    ,
                                    dataGridView1.Rows[23].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.8@@"
                                    ,
                                    dataGridView1.Rows[24].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.9@@"
                                    ,
                                    dataGridView1.Rows[25].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.10@@"
                                    ,
                                    dataGridView1.Rows[26].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.11@@"
                                    ,
                                    dataGridView1.Rows[27].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.12@@"
                                    ,
                                    dataGridView1.Rows[28].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.13@@"
                                    ,
                                    dataGridView1.Rows[29].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.14@@"
                                    ,
                                    dataGridView1.Rows[30].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.15@@"
                                    ,
                                    dataGridView1.Rows[31].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.16@@"
                                    ,
                                    dataGridView1.Rows[32].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.17@@"
                                    ,
                                    dataGridView1.Rows[33].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.18@@"
                                    ,
                                    dataGridView1.Rows[34].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.19@@"
                                    ,
                                    dataGridView1.Rows[35].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.20@@"
                                    ,
                                    dataGridView1.Rows[36].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.21@@"
                                    ,
                                    dataGridView1.Rows[37].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.22@@"
                                    ,
                                    dataGridView1.Rows[38].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.23@@"
                                    ,
                                    dataGridView1.Rows[39].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.24@@"
                                    ,
                                    dataGridView1.Rows[40].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.25@@"
                                    ,
                                    dataGridView1.Rows[41].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.26@@"
                                    ,
                                    dataGridView1.Rows[42].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.27@@"
                                    ,
                                    dataGridView1.Rows[43].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.28@@"
                                    ,
                                    dataGridView1.Rows[44].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.29@@"
                                    ,
                                    dataGridView1.Rows[45].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.30@@"
                                    ,
                                    dataGridView1.Rows[46].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.31@@"
                                    ,
                                    dataGridView1.Rows[47].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.32@@"
                                    ,
                                    dataGridView1.Rows[48].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.33@@"
                                    ,
                                    dataGridView1.Rows[49].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.34@@"
                                    ,
                                    dataGridView1.Rows[50].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.35@@",
                                    dataGridView1.Rows[51].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.36@@",
                                    dataGridView1.Rows[52].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.37@@",
                                    dataGridView1.Rows[53].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.38@@",
                                    dataGridView1.Rows[54].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.39@@",
                                    dataGridView1.Rows[55].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.40@@",
                                    dataGridView1.Rows[56].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@2.1.2.41@@",
                                    dataGridView1.Rows[57].Cells[1].Value.ToString()
                        )
                        ;

                    ReplaceTextWord(ref wdApp, "@@a1.1@@", dirtGridAnalogs.Rows[0].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.2@@", dirtGridAnalogs.Rows[1].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.3@@", dirtGridAnalogs.Rows[2].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.4@@", dirtGridAnalogs.Rows[3].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.5@@", dirtGridAnalogs.Rows[4].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.6@@", dirtGridAnalogs.Rows[5].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.7@@", dirtGridAnalogs.Rows[6].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.8@@", dirtGridAnalogs.Rows[7].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.9@@", dirtGridAnalogs.Rows[8].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.10@@", dirtGridAnalogs.Rows[9].Cells[1].Value.ToString());

                    ReplaceTextWord(ref wdApp, "@@a2.1@@", dirtGridAnalogs.Rows[0].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.2@@", dirtGridAnalogs.Rows[1].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.3@@", dirtGridAnalogs.Rows[2].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.4@@", dirtGridAnalogs.Rows[3].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.5@@", dirtGridAnalogs.Rows[4].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.6@@", dirtGridAnalogs.Rows[5].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.7@@", dirtGridAnalogs.Rows[6].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.8@@", dirtGridAnalogs.Rows[7].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.9@@", dirtGridAnalogs.Rows[8].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.10@@", dirtGridAnalogs.Rows[9].Cells[2].Value.ToString());

                    ReplaceTextWord(ref wdApp, "@@a3.1@@", dirtGridAnalogs.Rows[0].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.2@@", dirtGridAnalogs.Rows[1].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.3@@", dirtGridAnalogs.Rows[2].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.4@@", dirtGridAnalogs.Rows[3].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.5@@", dirtGridAnalogs.Rows[4].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.6@@", dirtGridAnalogs.Rows[5].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.7@@", dirtGridAnalogs.Rows[6].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.8@@", dirtGridAnalogs.Rows[7].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.9@@", dirtGridAnalogs.Rows[8].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.10@@", dirtGridAnalogs.Rows[9].Cells[3].Value.ToString());

                    ReplaceTextWord(ref wdApp, "@@b1.1@@", ((double)(houseCalcGrid.Rows[0].Cells[2].Value)).ToString("N", nfi) ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.2@@", ((double)(houseCalcGrid.Rows[1].Cells[2].Value)).ToString("N", nfi) ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.3@@", ((double)(houseCalcGrid.Rows[2].Cells[2].Value)).ToString("N", nfi) ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.4@@", houseCalcGrid.Rows[3].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.5@@", houseCalcGrid.Rows[4].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.6@@", houseCalcGrid.Rows[5].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.7@@", houseCalcGrid.Rows[6].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.8@@", houseCalcGrid.Rows[7].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.9@@", houseCalcGrid.Rows[8].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.10@@", houseCalcGrid.Rows[9].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.11@@", houseCalcGrid.Rows[10].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.12@@", houseCalcGrid.Rows[11].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.13@@", houseCalcGrid.Rows[12].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.14@@", houseCalcGrid.Rows[13].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.15@@", houseCalcGrid.Rows[14].Cells[2].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b1.16@@", houseCalcGrid.Rows[15].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.17@@", houseCalcGrid.Rows[16].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.18@@", houseCalcGrid.Rows[17].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.19@@", houseCalcGrid.Rows[18].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.20@@", houseCalcGrid.Rows[21].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.21@@", houseCalcGrid.Rows[20].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.22@@", houseCalcGrid.Rows[21].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.23@@", houseCalcGrid.Rows[22].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.24@@", houseCalcGrid.Rows[23].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.25@@", houseCalcGrid.Rows[24].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.26@@", houseCalcGrid.Rows[25].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.27@@", houseCalcGrid.Rows[26].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.28@@", houseCalcGrid.Rows[27].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.29@@", houseCalcGrid.Rows[28].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.30@@", houseCalcGrid.Rows[29].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.31@@", houseCalcGrid.Rows[30].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.32@@", houseCalcGrid.Rows[31].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.33@@", houseCalcGrid.Rows[32].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.34@@", houseCalcGrid.Rows[33].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b1.35@@", houseCalcGrid.Rows[34].Cells[2].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.1@@", ((double)(houseCalcGrid.Rows[0].Cells[3].Value)).ToString("N", nfi) ) ;  ReplaceTextWord(ref wdApp, "@@b2.2@@", ((double)(houseCalcGrid.Rows[1].Cells[3].Value)).ToString("N", nfi) ) ;  ReplaceTextWord(ref wdApp, "@@b2.3@@", ((double)(houseCalcGrid.Rows[2].Cells[3].Value)).ToString("N", nfi) ) ;  ReplaceTextWord(ref wdApp, "@@b2.4@@", houseCalcGrid.Rows[3].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.5@@", houseCalcGrid.Rows[4].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.6@@", houseCalcGrid.Rows[5].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.7@@", houseCalcGrid.Rows[6].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.8@@", houseCalcGrid.Rows[7].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.9@@", houseCalcGrid.Rows[8].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.10@@", houseCalcGrid.Rows[9].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.11@@", houseCalcGrid.Rows[10].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.12@@", houseCalcGrid.Rows[11].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.13@@", houseCalcGrid.Rows[12].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.14@@", houseCalcGrid.Rows[13].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.15@@", houseCalcGrid.Rows[14].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.16@@", houseCalcGrid.Rows[15].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.17@@", houseCalcGrid.Rows[16].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.18@@", houseCalcGrid.Rows[17].Cells[3].Value.ToString() ) ;  ReplaceTextWord(ref wdApp, "@@b2.19@@", houseCalcGrid.Rows[18].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.20@@", houseCalcGrid.Rows[21].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.21@@", houseCalcGrid.Rows[20].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.22@@", houseCalcGrid.Rows[21].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.23@@", houseCalcGrid.Rows[22].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.24@@", houseCalcGrid.Rows[23].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.25@@", houseCalcGrid.Rows[24].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.26@@", houseCalcGrid.Rows[25].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.27@@", houseCalcGrid.Rows[26].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.28@@", houseCalcGrid.Rows[27].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.29@@", houseCalcGrid.Rows[28].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.30@@", houseCalcGrid.Rows[29].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.31@@", houseCalcGrid.Rows[30].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.32@@", houseCalcGrid.Rows[31].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.33@@", houseCalcGrid.Rows[32].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.34@@", houseCalcGrid.Rows[33].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b2.35@@", houseCalcGrid.Rows[34].Cells[3].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b3.1@@", ((double)(houseCalcGrid.Rows[0].Cells[4].Value)).ToString("N", nfi) ) ; 
                    ReplaceTextWord(ref wdApp, "@@b3.2@@", ((double)(houseCalcGrid.Rows[1].Cells[4].Value)).ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b3.3@@", ((double)(houseCalcGrid.Rows[2].Cells[4].Value)).ToString("N", nfi) ) ; 
                    ReplaceTextWord(ref wdApp, "@@b3.4@@", houseCalcGrid.Rows[3].Cells[4].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b3.5@@", houseCalcGrid.Rows[4].Cells[4].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b3.6@@", houseCalcGrid.Rows[5].Cells[4].Value.ToString() ) ; 
                    ReplaceTextWord(ref wdApp, "@@b3.7@@", houseCalcGrid.Rows[6].Cells[4].Value.ToString() ) ;
                    ReplaceTextWord(ref wdApp, "@@b3.8@@", houseCalcGrid.Rows[7].Cells[4].Value.ToString() ) ;
                    ReplaceTextWord(ref wdApp, "@@b3.9@@", houseCalcGrid.Rows[8].Cells[4].Value.ToString() ) ;
                    ReplaceTextWord(ref wdApp, "@@b3.10@@", houseCalcGrid.Rows[9].Cells[4].Value.ToString() ) ;
                    ReplaceTextWord(ref wdApp, "@@b3.11@@", houseCalcGrid.Rows[10].Cells[4].Value.ToString() ) ;
                    ReplaceTextWord(ref wdApp, "@@b3.12@@", houseCalcGrid.Rows[11].Cells[4].Value.ToString() ) ;
                    ReplaceTextWord(ref wdApp, "@@b3.13@@", houseCalcGrid.Rows[12].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.14@@", houseCalcGrid.Rows[13].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.15@@", houseCalcGrid.Rows[14].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.16@@", houseCalcGrid.Rows[15].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.17@@", houseCalcGrid.Rows[16].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.18@@", houseCalcGrid.Rows[17].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.19@@", houseCalcGrid.Rows[18].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.20@@", houseCalcGrid.Rows[21].Cells[4].Value.ToString() ) ;  
                    ReplaceTextWord(ref wdApp, "@@b3.21@@", houseCalcGrid.Rows[20].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.22@@", houseCalcGrid.Rows[21].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.23@@", houseCalcGrid.Rows[22].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.24@@", houseCalcGrid.Rows[23].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.25@@", houseCalcGrid.Rows[24].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.26@@", houseCalcGrid.Rows[25].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.27@@", houseCalcGrid.Rows[26].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.28@@", houseCalcGrid.Rows[27].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.29@@", houseCalcGrid.Rows[28].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.30@@", houseCalcGrid.Rows[29].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.31@@", houseCalcGrid.Rows[30].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.32@@", houseCalcGrid.Rows[31].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.33@@", houseCalcGrid.Rows[32].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.34@@", houseCalcGrid.Rows[33].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.35@@", houseCalcGrid.Rows[34].Cells[4].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b1.36@@", ((double)(houseCalcGrid.Rows[35].Cells[2].Value)).ToString("N", nfi) ) ;

                    ReplaceTextWord(ref wdApp, "@@b2.36@@", ((double)(houseCalcGrid.Rows[35].Cells[3].Value)).ToString("N", nfi) ) ;

                    ReplaceTextWord(ref wdApp, "@@b3.36@@", ((double)(houseCalcGrid.Rows[35].Cells[4].Value)).ToString("N", nfi) ) ;

                    ReplaceTextWord(ref wdApp, "@@b4.1@@", houseCalcGrid.Rows[36].Cells[2].Value.ToString() ) ;

                    ReplaceTextWord(ref wdApp, "@@b4.2@@", ((double)(houseCalcGrid.Rows[37].Cells[2].Value)).ToString("N", nfi) ) ;

                    ReplaceTextWord(ref wdApp, "@@b4.3@@", ((double)(houseCalcGrid.Rows[38].Cells[2].Value)).ToString("N", nfi) ) ;

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "м2";
                    while (wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
                    {
                        wdApp.Selection.Characters[2].Font.Superscript = 1;
                    }
                    string te = wdApp.Selection.Text;

                    //saving
                    try
                    {
                        int x = wdDoc.Shapes.Count;
                        x = wdDoc.Shapes.Count;
                        for (int k = 1; k < x; k++)
                        {
                            Shape shape = wdDoc.Shapes[k];

                            if (shape.AlternativeText.Contains("cont"))
                            {
                                wdDoc.Shapes[k].TextEffect.Text = "№ " + contractNum.Text + " от " +
                                                                  calculationDate.Text + "г.";
                            }
                        }
                    }
                    catch (Exception exp)
                    {
                    }

                    wdApp.ActiveDocument.SaveAs(saveFileDialog1.FileName);

                    wdApp.Quit();
                }
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message);
                wdApp.Quit();
            }
        }

        private void AddGridCost(int dirtAnalogsCount, int columnNumber)
        {
            string analog0 = "@@ag" + columnNumber.ToString() + ".";
            string analogNum = "";
            string text = "";
            for (int i = 0; i < dirtAnalogsCount; i++)
            {
                text = dirtCalcGrid.Rows[i].Cells[columnNumber].Value.ToString();
                analogNum = analog0 + (i + 1).ToString() + "@@";
                ReplaceTextWord(ref wdApp, analogNum, text);
            }
        }

        private void AddHouseAnalog(int count, int columnNumber)
        {
            string analog0 = "@@d" + columnNumber.ToString() + ".";
            for (int i = 0; i < count; i++)
            {
                ReplaceTextWord(ref wdApp, analog0 + (i + 1).ToString() + "@@",
                                houseAnalogs.Rows[i].Cells[columnNumber + 1].Value.ToString());
            }
        }

        private void saveAppartmentsCalc_Click(object sender, EventArgs e)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();

            //Microsoft.Office.Interop.Excel.Workbook excelDoc = new Microsoft.Office.Interop.Excel.Workbook();
            string v = excelApp.Version;

            //excelApp.;
            string ownerFullName = ownerSurname.Text + " " + ownerName.Text + " " + ownerInit.Text;
            string customerFullName = customerSurname.Text + " " + customerName.Text + " " + customerInit.Text;
            string fileName = "отчет " + contractNum.Text + " расчет стоимости квартиры " + appartmentNum.Text + " " +
                              street.Text + " " + houseNum.Text + " для " + bankName.Text;
            saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();

            if (DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                excelApp.Workbooks.Open(System.Windows.Forms.Application.StartupPath + "\\calc.xls", Missing, Missing,
                                        Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
                                        Missing, Missing, Missing);

                //первый аналог
                excelApp.Workbooks[1].Sheets[1].Cells[2, 3] = calculationAppartaments[2, 0].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[2, 4] = calculationAppartaments[3, 0].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[2, 5] = calculationAppartaments[4, 0].Value;

                //                excelApp.Workbooks[1].Sheets[1].Cells[2, 6] = Convert.ToDateTime(analogsGrid.Rows[18].Cells[3].Value.ToString()).ToLongDateString();
                excelApp.Workbooks[1].Sheets[1].Cells[3, 3] = calculationAppartaments[2, 1].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[3, 4] = calculationAppartaments[3, 1].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[3, 5] = calculationAppartaments[4, 1].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[5, 3] = calculationAppartaments[2, 3].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[5, 4] = calculationAppartaments[3, 3].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[5, 5] = calculationAppartaments[4, 3].Value;
                string pattern = "MMMM yyyyг.";
                string d1 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[2].Value.ToString()).ToString(pattern);
                string d2 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[3].Value.ToString()).ToString(pattern);
                string d3 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[4].Value.ToString()).ToString(pattern);
                excelApp.Workbooks[1].Sheets[1].Cells[7, 3] = d1;
                excelApp.Workbooks[1].Sheets[1].Cells[7, 4] = d2;
                excelApp.Workbooks[1].Sheets[1].Cells[7, 5] = d3;

                excelApp.Workbooks[1].Sheets[1].Cells[8, 3] = calculationAppartaments[2, 6].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[8, 4] = calculationAppartaments[3, 6].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[8, 5] = calculationAppartaments[4, 6].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[10, 3] = calculationAppartaments[2, 8].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[10, 4] = calculationAppartaments[3, 8].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[10, 5] = calculationAppartaments[4, 8].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[12, 3] = calculationAppartaments[2, 10].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[12, 4] = calculationAppartaments[3, 10].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[12, 5] = calculationAppartaments[4, 10].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[14, 3] = calculationAppartaments[2, 12].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[14, 4] = calculationAppartaments[3, 12].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[14, 5] = calculationAppartaments[4, 12].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[16, 3] = calculationAppartaments[2, 14].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[16, 4] = calculationAppartaments[3, 14].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[16, 5] = calculationAppartaments[4, 14].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[18, 3] = calculationAppartaments[2, 16].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[18, 4] = calculationAppartaments[3, 16].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[18, 5] = calculationAppartaments[4, 16].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[20, 3] = calculationAppartaments[2, 18].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[20, 4] = calculationAppartaments[3, 18].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[20, 5] = calculationAppartaments[4, 18].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[22, 3] = calculationAppartaments[2, 20].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[22, 4] = calculationAppartaments[3, 20].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[22, 5] = calculationAppartaments[4, 20].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[24, 3] = calculationAppartaments[2, 22].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[24, 4] = calculationAppartaments[3, 22].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[24, 5] = calculationAppartaments[4, 22].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[26, 3] = calculationAppartaments[2, 24].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[26, 4] = calculationAppartaments[3, 24].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[26, 5] = calculationAppartaments[4, 24].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[28, 3] = calculationAppartaments[2, 26].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[28, 4] = calculationAppartaments[3, 26].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[28, 5] = calculationAppartaments[4, 26].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[30, 3] = calculationAppartaments[2, 28].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[30, 4] = calculationAppartaments[3, 28].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[30, 5] = calculationAppartaments[4, 28].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[32, 3] = calculationAppartaments[2, 30].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[32, 4] = calculationAppartaments[3, 30].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[32, 5] = calculationAppartaments[4, 30].Value;

                excelApp.Workbooks[1].Sheets[1].Cells[34, 3] = calculationAppartaments[2, 32].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[34, 4] = calculationAppartaments[3, 32].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[34, 5] = calculationAppartaments[4, 32].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[35, 3] = calculationAppartaments[2, 33].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[35, 4] = calculationAppartaments[3, 33].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[35, 5] = calculationAppartaments[4, 33].Value;
                excelApp.Workbooks[1].Sheets[1].Cells[38, 3] = calculationAppartaments[2, 36].Value;

                excelApp.ActiveWorkbook.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing,
                                               XlSaveAsAccessMode.xlNoChange,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelApp.Quit();
            }
        }

        private void calculationAppartaments_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            calculateCost();
        }

        private void calculationAppartaments_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            calculateCost();
        }

        private void dataGridView5_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            calculateCostHouse();
        }

        private void dataGridView5_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            calculateCostHouse();
        }

        public void calculateCostDirt()
        {
            try
            {
                string cellValue;
                cost_count1 = 0;
                cost_count2 = 0;
                cost_count3 = 0;
                cellValue = dirtCalcGrid.Rows[1].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cost1 = double.Parse(cellValue);
                }

                cellValue = dirtCalcGrid.Rows[2].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    m1 = double.Parse(cellValue);
                }
                cost_m1 = Math.Round(cost1 / m1);
                dirtCalcGrid.Rows[3].Cells[1].Value = cost_m1;

                cellValue = dirtCalcGrid.Rows[4].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost11 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[6].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost12 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[8].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost13 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[10].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost14 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[12].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost15 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[14].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost16 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[16].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost17 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[18].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost18 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[20].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost19 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[22].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost110 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cellValue = dirtCalcGrid.Rows[24].Cells[1].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost111 = double.Parse(cellValue);
                    cost_count1 = setCoefsCount(cellValue, cost_count1);
                }

                cor_cost_final1 =
                    Math.Round(cost_m1 * cor_cost11 * cor_cost12 * cor_cost13 * cor_cost14 * cor_cost15 * cor_cost16 * cor_cost17 *
                               cor_cost18 * cor_cost19 * cor_cost110 * cor_cost111);

                //второй аналог

                cellValue = dirtCalcGrid.Rows[1].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cost2 = double.Parse(cellValue);
                }

                cellValue = dirtCalcGrid.Rows[2].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    m2 = double.Parse(cellValue);
                }
                cost_m2 = Math.Round(cost2 / m2);
                dirtCalcGrid.Rows[3].Cells[2].Value = cost_m2;

                cellValue = dirtCalcGrid.Rows[4].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost21 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[6].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost22 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[8].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost23 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[10].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost24 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[12].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost25 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[14].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost26 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[16].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost27 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[18].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost28 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[20].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost29 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[22].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost210 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cellValue = dirtCalcGrid.Rows[24].Cells[2].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost211 = double.Parse(cellValue);
                    cost_count2 = setCoefsCount(cellValue, cost_count2);
                }

                cor_cost_final2 =
                    Math.Round(cost_m2 * cor_cost21 * cor_cost22 * cor_cost23 * cor_cost24 * cor_cost25 * cor_cost26 * cor_cost27 *
                               cor_cost28 * cor_cost29 * cor_cost210 * cor_cost211);

                //третий аналог

                cellValue = dirtCalcGrid.Rows[1].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cost3 = double.Parse(cellValue);
                }

                cellValue = dirtCalcGrid.Rows[2].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    m3 = double.Parse(cellValue);
                }
                cost_m3 = Math.Round(cost3 / m3);
                dirtCalcGrid.Rows[3].Cells[3].Value = cost_m3;

                cellValue = dirtCalcGrid.Rows[4].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost31 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[6].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost32 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[8].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost33 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[10].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost34 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[12].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost35 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[14].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost36 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[16].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost37 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[18].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost38 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[20].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost39 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[22].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost310 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cellValue = dirtCalcGrid.Rows[24].Cells[3].Value.ToString();
                if (cellValue != "")
                {
                    cor_cost311 = double.Parse(cellValue);
                    cost_count3 = setCoefsCount(cellValue, cost_count3);
                }

                cor_cost_final3 =
                    Math.Round(cost_m3 * cor_cost31 * cor_cost32 * cor_cost33 * cor_cost34 * cor_cost35 * cor_cost36 * cor_cost37 *
                               cor_cost38 * cor_cost39 * cor_cost310 * cor_cost311);

                dirtCalcGrid.Rows[26].Cells[1].Value = cost_count1;
                dirtCalcGrid.Rows[25].Cells[1].Value = cor_cost_final1;
                dirtCalcGrid.Rows[26].Cells[2].Value = cost_count2;
                dirtCalcGrid.Rows[25].Cells[2].Value = cor_cost_final2;
                dirtCalcGrid.Rows[26].Cells[3].Value = cost_count3;
                dirtCalcGrid.Rows[25].Cells[3].Value = cor_cost_final3;
                cor_cost_final1 =
                    Math.Round(cor_cost_final1 * Double.Parse(dirtCalcGrid.Rows[27].Cells[1].Value.ToString()));
                cor_cost_final2 =
                    Math.Round(cor_cost_final2 * Double.Parse(dirtCalcGrid.Rows[27].Cells[2].Value.ToString()));
                cor_cost_final3 =
                    Math.Round(cor_cost_final3 * Double.Parse(dirtCalcGrid.Rows[27].Cells[3].Value.ToString()));
                dirtCalcGrid.Rows[28].Cells[1].Value = cor_cost_final1;
                dirtCalcGrid.Rows[28].Cells[2].Value = cor_cost_final2;
                dirtCalcGrid.Rows[28].Cells[3].Value = cor_cost_final3;
                dirtCalcGrid.Rows[29].Cells[1].Value = Math.Round(cor_cost_final1 + cor_cost_final2 + cor_cost_final3);
                if (houseCalcGrid.Rows.Count > 0)
                {
                    houseCalcGrid.Rows[2].Cells[2].Value =
                        Math.Round(Double.Parse(dirtCalcGrid.Rows[29].Cells[1].Value.ToString()) *
                                   Double.Parse(houseCalcGrid.Rows[1].Cells[2].Value.ToString()) / 1000) * 1000;
                    houseCalcGrid.Rows[2].Cells[3].Value =
                        Math.Round(Double.Parse(dirtCalcGrid.Rows[29].Cells[1].Value.ToString()) *
                                   Double.Parse(houseCalcGrid.Rows[1].Cells[3].Value.ToString()) / 1000) * 1000;
                    houseCalcGrid.Rows[2].Cells[4].Value =
                        Math.Round(Double.Parse(dirtCalcGrid.Rows[29].Cells[1].Value.ToString()) *
                                   Double.Parse(houseCalcGrid.Rows[1].Cells[4].Value.ToString()) / 1000) * 1000;
                }
                dirtCalcGrid.Rows[29].Cells[2].Value = "";
                dirtCalcGrid.Rows[29].Cells[3].Value = "";
                final_cost_m =
                    Math.Round(Double.Parse(dirtCalcGrid.Rows[30].Cells[1].Value.ToString()) *
                               Double.Parse(dirtCalcGrid.Rows[29].Cells[1].Value.ToString()));
                dirtCalcGrid.Rows[31].Cells[1].Value = final_cost_m.ToString();
                dirtCalcGrid.Rows[32].Cells[1].Value = Math.Round(final_cost_m / 1000).ToString();
                dirtCalcGrid.Rows[33].Cells[1].Value =
                    Math.Round(Double.Parse(dirtCalcGrid.Rows[32].Cells[1].Value.ToString()) * 0.66).ToString();
                finalDirtCost = final_cost_m;
                likvidCostDirt = Math.Round(final_cost_m * 0.66);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void dirtCalcGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (!dirtCalcGrid.SelectedCells[0].Value.ToString().Contains('.'))
            {
                calculateCostDirt();
            }
        }

        private void dirtCalcGrid_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (!dirtCalcGrid.SelectedCells[0].Value.ToString().Contains('.'))
            {
                calculateCostDirt();
            }
        }

        private void town_TextChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void street_TextChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void houseNum_TextChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void buildingNum_TextChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void appartmentNum_TextChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void MO_TextChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void houseType_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void registrationDoc_TextChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        private void lift_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateTables();
        }

        public void addAtributeToXml(XmlTextWriter t, string name, string text)
        {
            if (text == null)
            {
                text = "";
            }
            t.WriteStartAttribute(name);
            t.WriteString(text);
            t.WriteEndAttribute();
        }

        /// <summary>
        ///     Save current state
        /// </summary>
        public void saveState()
        {
            string fileName = System.Windows.Forms.Application.StartupPath + "\\" + contractNum.Text + ".xml";

            saveXML(fileName);
            fileName = System.Windows.Forms.Application.StartupPath + "\\calcState.xml";
            File.Delete(fileName);
            saveXML(fileName);
        }

        /// <summary>
        ///     Load state from ini file
        /// </summary>
        public void loadState(string fileName)
        {
            var f = new FileStream(fileName, FileMode.OpenOrCreate);
            try
            {
                var settings = new XmlTextReader(f);
                while (settings.Read())
                {
                    if (settings.NodeType == XmlNodeType.Element)
                    {
                        if (settings.Name.Equals("test"))
                        {
                            customerName.Text = settings.GetAttribute(customerName.Name);
                            ownerName.Text = settings.GetAttribute(ownerName.Name);
                            customerInit.Text = settings.GetAttribute(customerInit.Name);
                            customerAddres.Text = settings.GetAttribute(customerAddres.Name);
                            customerPassDate.Text = settings.GetAttribute(customerPassDate.Name);
                            customerPassNum.Text = settings.GetAttribute(customerPassNum.Name);
                            houseNum.Text = settings.GetAttribute(houseNum.Name);
                            customerPassOVD.Text = settings.GetAttribute(customerPassOVD.Name);
                            customerPassport.Text = settings.GetAttribute(customerPassport.Name);
                            customerPhone.Text = settings.GetAttribute(customerPhone.Name);
                            customerSurname.Text = settings.GetAttribute(customerSurname.Name);
                            ownerName.Text = settings.GetAttribute(ownerName.Name);
                            ownerAddress.Text = settings.GetAttribute(ownerAddress.Name);
                            ownerInit.Text = settings.GetAttribute(ownerInit.Name);
                            ownerPassDate.Text = settings.GetAttribute(ownerPassDate.Name);
                            ownerPassNum.Text = settings.GetAttribute(ownerPassNum.Name);
                            ownerPassOVD.Text = settings.GetAttribute(ownerPassOVD.Name);
                            ownerPassport.Text = settings.GetAttribute(ownerPassport.Name);
                            ownerSurname.Text = settings.GetAttribute(ownerSurname.Name);
                            town.Text = settings.GetAttribute(town.Name);
                            street.Text = settings.GetAttribute(street.Name);
                            buildingNum.Text = settings.GetAttribute(buildingNum.Name);
                            roomsNum.Text = settings.GetAttribute(roomsNum.Name);
                            lm2text.Text = settings.GetAttribute(lm2text.Name);
                            m2text.Text = settings.GetAttribute(m2text.Name);
                            appartmentNum.Text = settings.GetAttribute(appartmentNum.Name);
                            calculationDate.Text = settings.GetAttribute(calculationDate.Name);
                            contractDate.Text = settings.GetAttribute(contractDate.Name);
                            contractNum.Text = settings.GetAttribute(contractNum.Name);
                            floors.Text = settings.GetAttribute(floors.Name);
                            floor.Text = settings.GetAttribute(floor.Name);
                            houseType.Text = settings.GetAttribute(houseType.Name);
                            registrationDoc.Text = settings.GetAttribute(registrationDoc.Name);
                            MO.Text = settings.GetAttribute(MO.Name);
                            string fname = fileName.Substring(0, fileName.LastIndexOf("."));
                            DataTable test = getDataFromXLS(fname + "Calc.xls");
                            calculationAppartaments.DataSource = test;

                            //Костыль, надо переделать
                            calculationAppartaments.Columns.RemoveAt(0);
                            calculationAppartaments.Columns.RemoveAt(0);
                            calculationAppartaments.Columns.RemoveAt(0);
                            calculationAppartaments.Columns.RemoveAt(0);
                            calculationAppartaments.Columns.RemoveAt(0);

                            calculationAppartaments.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                            calculationAppartaments.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                            calculationAppartaments.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                            calculationAppartaments.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                            calculationAppartaments.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;

                            //calculateCost();
                            test = null;
                            test = getDataFromXLS(fname + "Analogs.xls");
                            analogsGrid.DataSource = test;

                            //Костыль, надо переделать
                            analogsGrid.Columns.RemoveAt(0);
                            analogsGrid.Columns.RemoveAt(0);
                            analogsGrid.Columns.RemoveAt(0);
                            analogsGrid.Columns.RemoveAt(0);
                            analogsGrid.Columns.RemoveAt(0);

                            //analogsGrid.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                            //analogsGrid.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                            //analogsGrid.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                            //analogsGrid.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;

                            if (docTypeT == "Квартира")
                            {
                                objectDataGrid.Rows[2].Cells[1].Value = settings.GetAttribute("data2.1.1.2");
                                objectDataGrid.Rows[3].Cells[1].Value = settings.GetAttribute("data2.1.1.3");
                                objectDataGrid.Rows[4].Cells[1].Value = settings.GetAttribute("data2.1.1.4");
                                objectDataGrid.Rows[5].Cells[1].Value = settings.GetAttribute("data2.1.1.5");
                                objectDataGrid.Rows[6].Cells[1].Value = settings.GetAttribute("data2.1.1.6");
                                objectDataGrid.Rows[7].Cells[1].Value = settings.GetAttribute("data2.1.1.7");
                                objectDataGrid.Rows[8].Cells[1].Value = settings.GetAttribute("data2.1.1.8");
                                objectDataGrid.Rows[9].Cells[1].Value = settings.GetAttribute("data2.1.1.9");
                                objectDataGrid.Rows[10].Cells[1].Value = settings.GetAttribute("data2.1.1.10");
                                objectDataGrid.Rows[11].Cells[1].Value = settings.GetAttribute("data2.1.1.11");
                                objectDataGrid.Rows[12].Cells[1].Value = settings.GetAttribute("data2.1.1.12");
                                objectDataGrid.Rows[13].Cells[1].Value = settings.GetAttribute("data2.1.1.13");
                                objectDataGrid.Rows[14].Cells[1].Value = settings.GetAttribute("data2.1.1.14");
                                objectDataGrid.Rows[15].Cells[1].Value = settings.GetAttribute("data2.1.1.15");
                                objectDataGrid.Rows[16].Cells[1].Value = settings.GetAttribute("data2.1.1.16");

                                objectDataGrid.Rows[18].Cells[1].Value = settings.GetAttribute("data2.1.2.1");
                                objectDataGrid.Rows[19].Cells[1].Value = settings.GetAttribute("data2.1.2.2");
                                objectDataGrid.Rows[20].Cells[1].Value = settings.GetAttribute("data2.1.2.3");
                                objectDataGrid.Rows[21].Cells[1].Value = settings.GetAttribute("data2.1.2.4");
                                objectDataGrid.Rows[22].Cells[1].Value = settings.GetAttribute("data2.1.2.5");
                                objectDataGrid.Rows[23].Cells[1].Value = settings.GetAttribute("data2.1.2.6");
                                objectDataGrid.Rows[24].Cells[1].Value = settings.GetAttribute("data2.1.2.7");
                                objectDataGrid.Rows[25].Cells[1].Value = settings.GetAttribute("data2.1.2.8");
                                objectDataGrid.Rows[26].Cells[1].Value = settings.GetAttribute("data2.1.2.9");
                                objectDataGrid.Rows[27].Cells[1].Value = settings.GetAttribute("data2.1.2.10");
                                objectDataGrid.Rows[28].Cells[1].Value = settings.GetAttribute("data2.1.2.11");
                                objectDataGrid.Rows[29].Cells[1].Value = settings.GetAttribute("data2.1.2.12");
                                objectDataGrid.Rows[30].Cells[1].Value = settings.GetAttribute("data2.1.2.13");
                                objectDataGrid.Rows[31].Cells[1].Value = settings.GetAttribute("data2.1.2.14");
                                objectDataGrid.Rows[32].Cells[1].Value = settings.GetAttribute("data2.1.2.15");
                                objectDataGrid.Rows[33].Cells[1].Value = settings.GetAttribute("data2.1.2.16");
                                objectDataGrid.Rows[34].Cells[1].Value = settings.GetAttribute("data2.1.2.17");
                                objectDataGrid.Rows[35].Cells[1].Value = settings.GetAttribute("data2.1.2.18");
                                objectDataGrid.Rows[36].Cells[1].Value = settings.GetAttribute("data2.1.2.19");
                                objectDataGrid.Rows[37].Cells[1].Value = settings.GetAttribute("data2.1.2.20");
                                objectDataGrid.Rows[38].Cells[1].Value = settings.GetAttribute("data2.1.2.21");
                                objectDataGrid.Rows[39].Cells[1].Value = settings.GetAttribute("data2.1.2.22");

                                objectDataGrid.Rows[41].Cells[1].Value = settings.GetAttribute("data2.1.3.1");
                                objectDataGrid.Rows[42].Cells[1].Value = settings.GetAttribute("data2.1.3.2");
                                objectDataGrid.Rows[43].Cells[1].Value = settings.GetAttribute("data2.1.3.3");
                                objectDataGrid.Rows[44].Cells[1].Value = settings.GetAttribute("data2.1.3.4");
                                objectDataGrid.Rows[45].Cells[1].Value = settings.GetAttribute("data2.1.3.5");
                                objectDataGrid.Rows[46].Cells[1].Value = settings.GetAttribute("data2.1.3.6");
                                objectDataGrid.Rows[47].Cells[1].Value = settings.GetAttribute("data2.1.3.7");
                                objectDataGrid.Rows[48].Cells[1].Value = settings.GetAttribute("data2.1.3.8");
                                objectDataGrid.Rows[49].Cells[1].Value = settings.GetAttribute("data2.1.3.9");
                                objectDataGrid.Rows[50].Cells[1].Value = settings.GetAttribute("data2.1.3.10");
                                objectDataGrid.Rows[51].Cells[1].Value = settings.GetAttribute("data2.1.3.11");
                                objectDataGrid.Rows[52].Cells[1].Value = settings.GetAttribute("data2.1.3.12");
                                objectDataGrid.Rows[53].Cells[1].Value = settings.GetAttribute("data2.1.3.13");
                                objectDataGrid.Rows[54].Cells[1].Value = settings.GetAttribute("data2.1.3.14");
                                objectDataGrid.Rows[55].Cells[1].Value = settings.GetAttribute("data2.1.3.15");
                                objectDataGrid.Rows[56].Cells[1].Value = settings.GetAttribute("data2.1.3.16");
                                objectDataGrid.Rows[57].Cells[1].Value = settings.GetAttribute("data2.1.3.17");
                                objectDataGrid.Rows[58].Cells[1].Value = settings.GetAttribute("data2.1.3.18");
                                objectDataGrid.Rows[59].Cells[1].Value = settings.GetAttribute("data2.1.3.19");
                                objectDataGrid.Rows[60].Cells[1].Value = settings.GetAttribute("data2.1.3.20");
                                objectDataGrid.Rows[61].Cells[1].Value = settings.GetAttribute("data2.1.3.21");
                                objectDataGrid.Rows[62].Cells[1].Value = settings.GetAttribute("data2.1.3.22");
                                objectDataGrid.Rows[63].Cells[1].Value = settings.GetAttribute("data2.1.3.23");
                                objectDataGrid.Rows[64].Cells[1].Value = settings.GetAttribute("data2.1.3.24");
                                objectDataGrid.Rows[65].Cells[1].Value = settings.GetAttribute("data2.1.3.25");
                                objectDataGrid.Rows[66].Cells[1].Value = settings.GetAttribute("data2.1.3.26");
                                objectDataGrid.Rows[67].Cells[1].Value = settings.GetAttribute("data2.1.3.27");
                                objectDataGrid.Rows[68].Cells[1].Value = settings.GetAttribute("data2.1.3.28");
                                objectDataGrid.Rows[69].Cells[1].Value = settings.GetAttribute("data2.1.3.29");
                                objectDataGrid.Rows[70].Cells[1].Value = settings.GetAttribute("data2.1.3.30");
                                objectDataGrid.Rows[71].Cells[1].Value = settings.GetAttribute("data2.1.3.31");
                                objectDataGrid.Rows[72].Cells[1].Value = settings.GetAttribute("data2.1.3.32");
                                objectDataGrid.Rows[73].Cells[1].Value = settings.GetAttribute("data2.1.3.33");
                                objectDataGrid.Rows[74].Cells[1].Value = settings.GetAttribute("data2.1.3.34");
                                objectDataGrid.Rows[75].Cells[1].Value = settings.GetAttribute("data2.1.3.35");
                                objectDataGrid.Rows[76].Cells[1].Value = settings.GetAttribute("data2.1.3.36");
                                objectDataGrid.Rows[77].Cells[1].Value = settings.GetAttribute("data2.1.3.37");
                                objectDataGrid.Rows[78].Cells[1].Value = settings.GetAttribute("data2.1.3.38");
                                objectDataGrid.Rows[79].Cells[1].Value = settings.GetAttribute("data2.1.3.39");
                            }
                            ownerDocs.Text = settings.GetAttribute(ownerDocs.Name);
                        }
                    }
                }
                settings.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            f.Close();
        }

        private void customerSurname_TextChanged(object sender, EventArgs e)
        {
            customerPadeg();
        }

        private void ownerSurname_TextChanged(object sender, EventArgs e)
        {
            ownerPadeg();
        }

        private void m2text_TextChanged(object sender, EventArgs e)
        {
            switch (docTypeT)
            {
                case "Квартира":
                    {
                        objectDataGrid.Rows[48].Cells[1].Value = m2text.Text;
                        objectDataGrid.Rows[54].Cells[1].Value = m2text.Text;
                        calculationAppartaments.Rows[36].Cells[2].Value = m2text.Text;
                        analogsGrid.Rows[7].Cells[1].Value = m2text.Text;
                    }
                    break;

                case "Домовладение":
                    {
                        dataGridView1.Rows[35].Cells[1].Value = m2text.Text;
                        dataGridView1.Rows[38].Cells[1].Value = m2text.Text;
                        houseCalcGrid.Rows[37].Cells[2].Value = m2text.Text;
                        houseAnalogs.Rows[6].Cells[1].Value = m2text.Text;
                    }
                    break;

                case "Земельный участок":
                    {
                        //objectDataGrid.Rows[1].Cells[1].Value = fullAddress();
                    }
                    break;

                case "Домовладение с земельным участком":
                    {
                        dataGridView1.Rows[35].Cells[1].Value = m2text.Text;
                        dataGridView1.Rows[38].Cells[1].Value = m2text.Text;
                        houseCalcGrid.Rows[37].Cells[2].Value = m2text.Text;
                        houseAnalogs.Rows[6].Cells[1].Value = m2text.Text;
                    }
                    break;

                default:
                    break;
            }
        }

        private void lm2text_TextChanged(object sender, EventArgs e)
        {
            switch (docTypeT)
            {
                case "Квартира":
                    {
                        objectDataGrid.Rows[49].Cells[1].Value = lm2text.Text;
                    }
                    break;

                case "Домовладение":
                    {
                        dataGridView1.Rows[36].Cells[1].Value = lm2text.Text;
                    }
                    break;

                case "Земельный участок":
                    {
                        //objectDataGrid.Rows[1].Cells[1].Value = fullAddress();
                    }
                    break;

                case "Домовладение с земельным участком":
                    {
                        dataGridView1.Rows[36].Cells[1].Value = lm2text.Text;
                    }
                    break;

                default:
                    break;
            }
        }

        private string CreateFileName()
        {
            string townName = " " + town.Text + ", ";

            if ((town.Text == "г. Владикавказ") || (town.Text == "г.Владикавказ"))
            {
                townName = " ";
            }

            string buildNum = null;

            if (buildingNum.Text != "")
            {
                buildNum = ", корп. " + buildingNum.Text;
            }
            roomsAsString();
            string fileName = "";
            if (bankName.Text == "втб 24")
            {
                fileName = "отчет  " + contractNum.Text + " от " + calculationDate.Text + "г " + roomsN + " квартира " +
                           appartmentNum.Text + " " + townName + " " + street.Text + " " + houseNum.Text + " " +
                           buildNum + " " + ownerSurname.Text + " " + ownerName.Text + " для " + customerSurname.Text +
                           " " + customerName.Text + " 24 втб";
            }

            else if (bankName.Text == "сбербанк")
            {
                fileName = "отчет  " + contractNum.Text + " от " + calculationDate.Text + "г " + roomsN + " квартира " +
                           appartmentNum.Text + townName + street.Text + ", " + houseNum.Text + buildNum + " " +
                           ownerSurname.Text + " " + ownerName.Text + " для " + customerSurname.Text + " " +
                           customerName.Text + " ипотека " + bankName.Text;
            }

            else if (bankName.Text == "брр")
            {
                fileName = "отчет  " + contractNum.Text + " от " + calculationDate.Text + "г " + roomsN + " квартира " +
                           appartmentNum.Text + townName + street.Text + ", " + houseNum.Text + buildNum + " " +
                           ownerSurname.Text + " " + ownerName.Text + " для " + customerSurname.Text + " " +
                           customerName.Text + bankName.Text;
            }
            else if (bankName.Text == "аижк")
            {
                fileName = "отчет  " + contractNum.Text + " от " + calculationDate.Text + "г " + roomsN + " квартира " +
                           appartmentNum.Text + townName + street.Text + ", " + houseNum.Text + buildNum + " " +
                           ownerSurname.Text + " " + ownerName.Text + " для " + customerSurname.Text + " " +
                           customerName.Text + " ипотека для" + bankName.Text;
            }
            else if (bankName.Text == "банк москвы")
            {
                fileName = "отчет  " + contractNum.Text + " от " + calculationDate.Text + "г " + roomsN + " квартира " +
                           appartmentNum.Text + " " + townName + " " + street.Text + ", " + houseNum.Text + buildNum +
                           " " + ownerSurname.Text + " " + ownerName.Text + " для " + customerSurname.Text + " " +
                           customerName.Text + " " + bankName.Text;
            }
            else
            {
                fileName = "отчет  " + contractNum.Text + " от " + calculationDate.Text + "г " + roomsN + " квартира " +
                           appartmentNum.Text + " " + townName + " " + street.Text + ", " + houseNum.Text + buildNum +
                           " " + ownerSurname.Text + " " + ownerName.Text + " для " + customerSurname.Text + " " +
                           customerName.Text + " " + bankName.Text;
            }

            fileName = fileName.Replace("\"", " ").ToLower();
            fileName = fileName.Replace("/", " ").ToLower();
            fileName = fileName.Replace(",", " ").ToLower();
            fileName = fileName.Replace("№", " ").ToLower();
            fileName = fileName.Replace(".", " ").ToLower();
            fileName = fileName.Replace("-", " ").ToLower();

            //saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();
            return fileName;
        }

        private void saveResultButton_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = CreateFileName();
            switch (houseType.Text)
            {
                case "Кирпичный":
                    {
                        houseType1 = "кирпичного";
                    }
                    break;

                case "Панельный":
                    {
                        houseType1 = "панельного";
                    }
                    break;

                case "Монолитный":
                    {
                        houseType1 = "монолитного";
                    }
                    break;

                default:
                    break;
            }

            try
            {
                if (DialogResult.OK == saveFileDialog1.ShowDialog())
                {
                    wdApp = new Application();
                    var wdDoc = new Document();

                    wdDoc = wdApp.Documents.Open(System.Windows.Forms.Application.StartupPath + "\\m2.doc", Missing, true);
                    wdApp.ActiveDocument.Words[1].Select();
                    wdApp.Selection.Copy();
                    wdDoc.Close();
                    string template = "\\шаблоны\\ОсновнойШаблон.doc";

                    if (bankName.Text == "втб 24")
                    {
                        template = "\\шаблоны\\ВТБ24.doc";
                    }

                    if (ownerOrg.Checked)
                    {
                        template = "\\шаблоны\\Организация.doc";
                    }

                    wdDoc = wdApp.Documents.Open(System.Windows.Forms.Application.StartupPath + template, Missing, true);
                    object replaceAll = WdReplace.wdReplaceAll;

                    // Gets a NumberFormatInfo associated with the en-US culture.
                    NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;

                    nfi.NumberDecimalDigits = 0;
                    nfi.NumberGroupSeparator = " ";

                    nfi.PositiveSign = "";
                    customerPadeg();

                    string ownerFullName = ownerSurname.Text + " " + ownerName.Text + " " + ownerInit.Text;
                    string customerFullName = customerSurname.Text + " " + customerName.Text + " " + customerInit.Text;

                    calculationDate.CustomFormat = "dd MMMM yyyy";
                    string calculationDateStr = calculationDate.Text;
                    int lenght = calculationDateStr.Length;
                    string temp = null;
                    string t;

                    for (int i = 0; i < lenght; i++)
                    {
                        if (i == 3)
                        {
                            t = calculationDateStr[i].ToString().ToUpper();
                            temp += t;
                        }
                        else
                        {
                            temp += calculationDateStr[i];
                        }
                    }

                    calculationDateStr = temp;

                    calculationDate.CustomFormat = "dd/MM/yy";
                    int sentencesCount = wdDoc.Sentences.Count;
                    string topColontitul = topColontitulCreator();

                    wdDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = topColontitul;

                    ReplaceTextWord(ref wdApp, "@@MO@@", MO.Text);

                    if (ownerOrg.Checked)
                    {
                        if (bankName.Text == "втб 24")
                        {
                            ReplaceTextWord(ref wdApp, "@@ownerOrgname@@",
                                            "Операционный офис «Владикавказский» филиала №2351 ВТБ 24 (ЗАО)");
                            ReplaceTextWord(ref wdApp, "@@INN@@", "7710353606");
                            ReplaceTextWord(ref wdApp, "@@OGRN@@", "1027739207462");
                            ReplaceTextWord(ref wdApp, "@@KPP@@", "231002001");
                            ReplaceTextWord(ref wdApp, "@@orgAddress@@", "РСО-Алания, г. Владикавказ, ул. Коцоева, д.13");
                        }
                        else
                        {
                            ReplaceTextWord(ref wdApp, "@@ownerOrgname@@", orgName.Text);
                            ReplaceTextWord(ref wdApp, "@@INN@@", orgINN.Text);
                            ReplaceTextWord(ref wdApp, "@@OGRN@@", orgOGRN.Text);
                            ReplaceTextWord(ref wdApp, "@@KPP@@", orgKPP.Text);
                            ReplaceTextWord(ref wdApp, "@@orgAddress@@", orgAdd.Text);
                        }
                    }

                    ReplaceTextWord(ref wdApp, "@@houseType1@@", houseType1);
                    ReplaceTextWord(ref wdApp, "@@calculationDateStr@@", calculationDateStr);
                    ReplaceTextWord(ref wdApp, "@@houseType@@", houseType.Text.ToLower());
                    ReplaceTextWord(ref wdApp, "@@roomsT@@", roomsT);
                    ReplaceTextWord(ref wdApp, "@@roomsX@@", roomsX);
                    ReplaceTextWord(ref wdApp, "@@lm2@@", lm2text.Text);
                    ReplaceTextWord(ref wdApp, "@@m2@@", m2text.Text);
                        ReplaceTextWord(ref wdApp, "@@raion@@", ", " + textBox1.Text);
                    ReplaceTextWord(ref wdApp, "@@customerNameInits@@", customerFamiliyR + " " + getInits());
                    ReplaceTextWord(ref wdApp, "@@calculationDate@@", calculationDate.Text);
                    if (newBuildingCheck.Checked)
                    {
                        var wdNew = new Document();

                        wdNew = wdApp.Documents.Open(
                            System.Windows.Forms.Application.StartupPath + "\\новостройка.doc", Missing, true);
                        wdNew.Sections[1].Range.Select();
                        wdNew.Sections[1].Range.Copy();
                        wdNew.Close();

                        wdApp.Selection.Find.ClearFormatting();
                        wdApp.Selection.Find.Text = "@@новостройка@@";
                        while (wdApp.Selection.Find.Execute(
                            ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                            ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                            ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
                        {
                            //wdNew = wdApp.Documents.Open(System.Windows.Forms.Application.StartupPath + "новостройка.doc", Missing, true);

                            //wdApp.Selection.Text = "";
                            wdApp.Selection.Paste();

                            wdApp.ActiveDocument.Sections[1].Range.Select();
                        }
                    }
                    else
                    {
                        ReplaceTextWord(ref wdApp, "@@новостройка@@", "");
                    }

                    ReplaceTextWord(ref wdApp, "@@customerFullname@@", customerFullName);

                    roomsAsString();
                    ReplaceTextWord(ref wdApp, "@@rooms1@@", rooms1);

                    ReplaceTextWord(ref wdApp, "@@customerFullnameR@@", customerFullNameR);
                    ReplaceTextWord(ref wdApp, "@@customerFullnameT@@", customerFullNameT);

                    ReplaceTextWord(ref wdApp, "@@customerFullnameD@@", customerFullNameD);
                    ReplaceTextWord(ref wdApp, "@@rooms@@", roomsAsString());
                    ReplaceTextWord(ref wdApp, "@@appartmentNum@@", "№" + appartmentNum.Text);
                    ReplaceTextWord(ref wdApp, "@@street@@", street.Text);
                    ReplaceTextWord(ref wdApp, "@@houseNum@@", houseNum.Text);

                    string buildNum = null;
                    if (buildingNum.Text != "")
                    {
                        buildNum = ", корп." + buildingNum.Text;
                    }
                    else
                    {
                        buildNum = buildingNum.Text;
                    }
                    ReplaceTextWord(ref wdApp, "@@buildingNum@@", buildNum);
                    ReplaceTextWord(ref wdApp, "@@customerAddress@@", customerAddres.Text);
                    ReplaceTextWord(ref wdApp, "@@floor@@", floor.Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@floors@@", floors.Text);
                    ReplaceTextWord(ref wdApp, "@@town@@", town.Text);
                    ReplaceTextWord(ref wdApp, "@@cost@@", finalCostRounded.ToString("N", nfi));
                    ReplaceTextWord(ref wdApp, "@@contractNum@@", contractNum.Text);
                    ReplaceTextWord(ref wdApp, "@@contractDate@@", contractDate.Text);
                    ReplaceTextWord(ref wdApp, "@@customerName@@", customerName.Text);
                    ReplaceTextWord(ref wdApp, "@@customerInit@@", customerInit.Text);
                    ReplaceTextWord(ref wdApp, "@@likvidCost@@", likvidCost.ToString("N", nfi));
                    ReplaceTextWord(ref wdApp, "@@stringCost@@", costStr.ToLower());

                    getUvaj();
                    ReplaceTextWord(ref wdApp, "@@uvaj@@", uvaj);

                    //Customer Passport
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerPassport@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerPassport.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerPassNum@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerPassNum.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerPassOVD@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerPassOVD.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerPassDate@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerPassDate.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerFullAddress@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerAddres.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    //owner Passport
                    if (owners.Count > 1)
                    {
                        int ownerIndex = 0;
                        foreach (Owner owner in owners)
                        {
                            ownerIndex++;
                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerFullnameD@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = owner.ownerFullNameD + "; @@ownerFullnameD@@";

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                            //ReplaceTextWord(ref wdApp, "@@ownerFullnameD@@",owner.ownerFullNameD + "; @@ownerFullnameD@@");

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerFullnameT@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = owner.ownerFullNameT + "; @@ownerFullnameT@@";

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                            //ReplaceTextWord(ref wdApp, "@@ownerFullnameT@@", owner.ownerFullNameT + "; @@ownerFullnameT@@");

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerFullnameR@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = owner.ownerFullNameR + "; @@ownerFullnameR@@";

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                            //ReplaceTextWord(ref wdApp, "@@ownerFullnameR@@", owner.ownerFullNameR + "; @@ownerFullnameR@@");

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerFullname1@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = owner.ownerFullName + ";/rn@@ownerFullname1@@";

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                            //ReplaceTextWord(ref wdApp, "@@ownerFullname1@@", owner.ownerFullName + ";/rn@@ownerFullname1@@");

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerFullname@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = ownerIndex + "." + owner.ownerFullName + "/rn";

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                            //ReplaceTextWord(ref wdApp, "@@ownerFullname@@", ownerIndex + "." + owner.ownerFullName + "/rn");

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@passportSerial@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = owner.passportSerial;

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                            //ReplaceTextWord(ref wdApp, "@@passportSerial@@", owner.passportSerial);

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerPassport@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = ownerPassport.Text;

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                            //ReplaceTextWord(ref wdApp, "@@ownerPassport@@", ownerPassport.Text);

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerPassNum@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = owner.passNum;

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                            //ReplaceTextWord(ref wdApp, "@@ownerPassNum@@", owner.passNum);

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerPassOVD@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = owner.OVD;

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                            //ReplaceTextWord(ref wdApp, "@@ownerPassOVD@@", owner.OVD);

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerPassDate@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = owner.passDate;

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                            //ReplaceTextWord(ref wdApp, "@@ownerPassDate@@", owner.passDate);

                            wdApp.Selection.Find.ClearFormatting();
                            wdApp.Selection.Find.Text = "@@ownerFullAddress@@";
                            wdApp.Selection.Find.Replacement.ClearFormatting();
                            wdApp.Selection.Find.Replacement.Text = owner.address + ";/rn" +
                                                                    "@@ownerFullname@@ Паспорт гражданина РФ серии @@ownerPassport@@ №@@ownerPassNum@@, выдан @@ownerPassDate@@ @@ownerPassOVD@@. Проживает по адресу: @@ownerFullAddress@@";

                            wdApp.Selection.Find.Execute(
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                            //ReplaceTextWord(ref wdApp, "@@ownerFullAddress@@", owner.address + ";/rn" + "@@ownerFullname@@ Паспорт гражданина РФ серии @@ownerPassport@@ №@@ownerPassNum@@, выдан @@ownerPassDate@@ @@ownerPassOVD@@. Проживает по адресу: @@ownerFullAddress@@");
                        }

                        ReplaceTextWord(ref wdApp, "; @@ownerFullnameD@@", "");
                        ReplaceTextWord(ref wdApp, "; @@ownerFullnameT@@", "");
                        ReplaceTextWord(ref wdApp, "; @@ownerFullnameR@@", "");
                        ReplaceTextWord(ref wdApp, "; @@ownerFullname@@", "");
                        ReplaceTextWord(ref wdApp, ";/rn@@ownerFullname1@@", "");
                        ReplaceTextWord(ref wdApp, "; @@ownerPassport@@", "");
                        ReplaceTextWord(ref wdApp, "; @@ownerPassNum@@", "");
                        ReplaceTextWord(ref wdApp, "; @@ownerPassOVD@@", "");
                        ReplaceTextWord(ref wdApp, "; @@ownerPassDate@@", "");
                        ReplaceTextWord(ref wdApp, "; @@ownerFullAddress@@", "");
                        ReplaceTextWord(ref wdApp,
                                        ";/rn@@ownerFullname@@ Паспорт гражданина РФ серии @@ownerPassport@@ №@@ownerPassNum@@, выдан @@ownerPassDate@@ @@ownerPassOVD@@. Проживает по адресу: @@ownerFullAddress@@",
                                        "");
                        InsertParagraphs(ref wdApp);
                    }
                    else
                    {
                        ReplaceTextWord(ref wdApp, "@@ownerFullnameD@@", ownerFullNameD);
                        ReplaceTextWord(ref wdApp, "@@ownerFullnameT@@", ownerFullNameT);
                        ReplaceTextWord(ref wdApp, "@@ownerFullnameR@@", ownerFullNameR);
                        ReplaceTextWord(ref wdApp, "@@ownerFullname@@", ownerFullName);
                        ReplaceTextWord(ref wdApp, "@@ownerFullname1@@", ownerFullName);
                        ReplaceTextWord(ref wdApp, "@@ownerPassport@@", ownerPassport.Text);
                        ReplaceTextWord(ref wdApp, "@@ownerPassNum@@", ownerPassNum.Text);
                        ReplaceTextWord(ref wdApp, "@@ownerPassOVD@@", ownerPassOVD.Text);
                        ReplaceTextWord(ref wdApp, "@@ownerPassDate@@", ownerPassDate.Text);
                        ReplaceTextWord(ref wdApp, "@@ownerFullAddress@@", ownerAddress.Text);
                    }

                    //
                    ReplaceTextWord(ref wdApp, "@@ownerDoc@@", ownerDocs.Text);
                    ReplaceTextWord(ref wdApp, "@@registrationDoc@@", registrationDoc.Text);
                    ReplaceTextWord(ref wdApp, "@@ownerFullnameD@@", ownerFullNameD);
                    ReplaceTextWord(ref wdApp, "@@ownerFullnameT@@", ownerFullNameT);
                    ReplaceTextWord(ref wdApp, "@@ownerFullnameR@@", ownerFullNameR);
                    ReplaceTextWord(ref wdApp, "@@ownerFullname@@", ownerFullName);
                    ReplaceTextWord(ref wdApp, "@@ownerFullname1@@", ownerFullName);
                    ReplaceTextWord(ref wdApp, "@@ownerPassport@@", ownerPassport.Text);
                    ReplaceTextWord(ref wdApp, "@@ownerPassNum@@", ownerPassNum.Text);
                    ReplaceTextWord(ref wdApp, "@@ownerPassOVD@@", ownerPassOVD.Text);
                    ReplaceTextWord(ref wdApp, "@@ownerPassDate@@", ownerPassDate.Text);
                    ReplaceTextWord(ref wdApp, "@@ownerFullAddress@@", ownerAddress.Text);

                    var padeg = new Declension();

                    string test = objectDataGrid.Rows[41].Cells[1].Value.ToString();
                    kadastr = padeg.GetAppointmentPadeg(test, 2);
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@tehPass@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = kadastr;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[2].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[3].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[4].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[5].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[6].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[7].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[8].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[9].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[10].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[11].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[12].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[13].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[14].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[15].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[16].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[18].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[19].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[20].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[21].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[22].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[23].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[24].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[25].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[26].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[27].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[28].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[29].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[30].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[31].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[32].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[33].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.17@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[34].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.18@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[35].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.19@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[36].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.20@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[37].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.21@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[38].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.22@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[39].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[41].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[42].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[43].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[44].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[45].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[46].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[47].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[48].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[49].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[50].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[51].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[52].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[53].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[54].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[55].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[56].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.17@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[57].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.18@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[58].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.19@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[59].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.20@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[60].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.21@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[61].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.22@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[62].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    /*wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.23@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[63].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);*/
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.23@@";
                    while (wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
                    {
                        //Microsoft.Office.Interop.Word.Range r1;

                        //r1.Text = objectDataGrid.Rows[63].Cells[1].Value.ToString();
                        wdApp.Selection.Text = objectDataGrid.Rows[63].Cells[1].Value.ToString();

                        //wdApp.Selection.Font.Superscript = 1;
                        //                        wdApp.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault);
                        wdApp.ActiveDocument.Sections[1].Range.Select();
                    }

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.24@@";
                    while (wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
                    {
                        //Microsoft.Office.Interop.Word.Range r1;

                        //r1.Text = objectDataGrid.Rows[63].Cells[1].Value.ToString();
                        wdApp.Selection.Text = objectDataGrid.Rows[64].Cells[1].Value.ToString();

                        //wdApp.Selection.Font.Superscript = 1;
                        //                        wdApp.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault);
                        wdApp.ActiveDocument.Sections[1].Range.Select();
                    }

                    //wdApp.Selection.Find.ClearFormatting();
                    //wdApp.Selection.Find.Text = "@@2.1.3.24@@";
                    //wdApp.Selection.Find.Replacement.ClearFormatting();
                    //wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[64].Cells[1].Value.ToString();

                    //wdApp.Selection.Find.Execute(
                    //             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    //             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    //             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.25@@";
                    while (wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
                    {
                        wdApp.Selection.Text = objectDataGrid.Rows[65].Cells[1].Value.ToString();

                        wdApp.ActiveDocument.Sections[1].Range.Select();
                    }

                    //wdApp.Selection.Find.ClearFormatting();
                    //wdApp.Selection.Find.Text = "@@2.1.3.25@@";
                    //wdApp.Selection.Find.Replacement.ClearFormatting();
                    //wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[65].Cells[1].Value.ToString();

                    //wdApp.Selection.Find.Execute(
                    //             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    //             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    //             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.26@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[66].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.27@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[67].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.28@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[68].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.29@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[69].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.30@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[70].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.31@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[71].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.32@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[72].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.33@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[73].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.34@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[74].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.35@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[75].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.36@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[76].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.37@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[77].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.38@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[78].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.3.39@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[79].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[0].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[1].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[2].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[3].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[4].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[5].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[6].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[7].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[8].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[9].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[10].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[11].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[12].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[13].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a0.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[14].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[0].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[1].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[2].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[3].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[4].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[5].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[6].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[7].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[8].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[9].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[10].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[11].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[12].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[13].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[14].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[15].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.17@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[16].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.18@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[17].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.19@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[18].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a1.20@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[19].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[0].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[1].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[2].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[3].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[4].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[5].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[6].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[7].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[8].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[9].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[10].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[11].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[12].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[13].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[14].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[15].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.17@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[16].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.18@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[17].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.19@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[18].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a2.20@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[19].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[0].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[1].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[2].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[3].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[4].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[5].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[6].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[7].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[8].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[9].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[10].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[11].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[12].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[13].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[14].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[15].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.17@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[16].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.18@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[17].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.19@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[18].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@a3.20@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = analogsGrid.Rows[19].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[0].Cells[2].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[1].Cells[2].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[2].Cells[2].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[3].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[4].Cells[2].Value.ToString();
                    string pattern = "MMMM yyyyг.";
                    string d1 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[2].Value.ToString()).ToString(pattern);
                    string d2 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[3].Value.ToString()).ToString(pattern);
                    string d3 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[4].Value.ToString()).ToString(pattern);
                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();

                    wdApp.Selection.Find.Replacement.Text = d1;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[6].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[7].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[8].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[9].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[10].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[11].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[12].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[13].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[14].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[15].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.17@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[16].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.18@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[17].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.19@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[18].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.20@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[21].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.21@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[20].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.22@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[21].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.23@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[22].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.24@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[23].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.25@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[24].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.26@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[25].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.27@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[26].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.28@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[27].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.29@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[28].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.30@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[29].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.31@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[30].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.32@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[31].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.33@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[32].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.34@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[33].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b1.35@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[34].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[0].Cells[3].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[1].Cells[3].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[2].Cells[3].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[3].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[4].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();

                    wdApp.Selection.Find.Replacement.Text = d2;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[6].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[7].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[8].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[9].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[10].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[11].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[12].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[13].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[14].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[15].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.17@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[16].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.18@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[17].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.19@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[18].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.20@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[21].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.21@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[20].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.22@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[21].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.23@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[22].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.24@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[23].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.25@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[24].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.26@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[25].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.27@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[26].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.28@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[27].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.29@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[28].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.30@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[29].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.31@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[30].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.32@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[31].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.33@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[32].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.34@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[33].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b2.35@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[34].Cells[3].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[0].Cells[4].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[1].Cells[4].Value)).ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[2].Cells[4].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[3].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[4].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();

                    wdApp.Selection.Find.Replacement.Text = d3;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[6].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[7].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[8].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[9].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[10].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[11].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[12].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[13].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[14].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[15].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.17@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[16].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.18@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[17].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.19@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[18].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.20@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[21].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.21@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[20].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.22@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[21].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.23@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[22].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.24@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[23].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.25@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[24].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.26@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[25].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.27@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[26].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.28@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[27].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.29@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[28].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.30@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[29].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.31@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[30].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.32@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[31].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.33@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[32].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.34@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[33].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b3.35@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[34].Cells[4].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b4.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[35].Cells[2].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b4.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[36].Cells[2].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b4.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[37].Cells[2].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@b4.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text =
                        ((double)(calculationAppartaments.Rows[38].Cells[2].Value)).ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "м2";
                    while (wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
                    {
                        wdApp.Selection.Characters[2].Font.Superscript = 1;
                    }

                    //saving
                    try
                    {
                        int x = wdDoc.Shapes.Count;
                        x = wdDoc.Shapes.Count;
                        for (int k = 1; k < x; k++)
                        {
                            Shape shape = wdDoc.Shapes[k];

                            //string l = shape.AlternativeText;
                            if (shape.AlternativeText.Contains("cont"))
                            {
                                wdDoc.Shapes[k].TextEffect.Text = "№ " + contractNum.Text + " от " +
                                                                  calculationDate.Text + "г.";
                            }
                        }
                        /*
                       for (int k = 1; k < x; k++)
                        {
                            Microsoft.Office.Interop.Word.Shape shape = wdDoc.Shapes[k];

                            if (shape.AlternativeText.Contains("first"))
                            {
                                System.Drawing.Image firstPageImg = System.Drawing.Image.FromFile(imagesGrid.Rows[0].Cells[2].Value.ToString());

                               //Clipboard.SetImage(firstPageImg);
                                shape.Select();
                                wdDoc.Shapes[k].CanvasItems.AddPicture(imagesGrid.Rows[0].Cells[2].Value.ToString());
                                wdDoc.Shapes[k].Apply();

                               // wdApp.Selection.PasteSpecial();
                               // wdApp.ActiveDocument.Shapes.AddPicture(imagesGrid.Rows[0].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing, Type.Missing, 500, 370, Type.Missing);
                                //wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[0].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                //wdApp.Selection.InlineShapes
                                //Clipboard.Clear();
                            }
                        }*/
                    }
                    catch (Exception exp)
                    {
                    }

                    /*foreach (Microsoft.Office.Interop.Word.Table table in wdApp.ActiveDocument.Tables)
                    {
                        try
                        {
                            //  if (table.Columns[0].Cells[0].Range.Text.Contains("@@1@@"))
                            //{
                            foreach (Microsoft.Office.Interop.Word.Column col in table.Columns)
                            {
                                foreach (Microsoft.Office.Interop.Word.Cell cell in col.Cells)
                                {
                                    int rowCount = imagesGrid.RowCount;
                                    string l = cell.Range.Text;
                                    if (l.Contains("@@1@@"))
                                    {
                                        cell.Select();
                                        cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[1].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@1@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }

                                    if (l.Contains("@@2@@"))
                                    {
                                        cell.Select();

                                        //cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[2].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@2@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }
                                    if (l.Contains("@@3@@"))
                                    {
                                        cell.Select();

                                        //cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[3].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@3@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }

                                    if (l.Contains("@@4@@"))
                                    {
                                        cell.Select();

                                        //cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[4].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@4@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }

                                    if (l.Contains("@@5@@"))
                                    {
                                        cell.Select();

                                        //cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[5].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@5@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }

                                    if (l.Contains("@@6@@"))
                                    {
                                        cell.Select();

                                        //cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[6].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@6@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }

                                    if (l.Contains("@@7@@"))
                                    {
                                        cell.Select();

                                        //cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[7].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@7@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }

                                    if (l.Contains("@@8@@"))
                                    {
                                        cell.Select();

                                        //cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[8].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@8@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }

                                    if (l.Contains("@@9@@"))
                                    {
                                        cell.Select();

                                        //cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[9].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@9@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }

                                    if (l.Contains("@@10@@"))
                                    {
                                        cell.Select();

                                        //cell.Range.Text = "";
                                        wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[10].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
                                        wdApp.Selection.Find.ClearFormatting();
                                        wdApp.Selection.Find.Text = "@@10@@";
                                        wdApp.Selection.Find.Replacement.ClearFormatting();
                                        wdApp.Selection.Find.Replacement.Text = "";

                                        wdApp.Selection.Find.Execute(
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                                     ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                                    }
                                }
                            }
                        }

                        //  }
                        catch (Exception except)
                        {
                        }
                    }

                    /*for (int k = 1; k < x; k++)
                    {
                        Microsoft.Office.Interop.Word.Shape shape = wdDoc.Shapes[k];
                        float shift = 150;
                        string l = shape.AlternativeText;
                        if (l == "facade")
                        {
                            shape.IncrementTop(-shift);
                        }
                        if (l == "appartmentNum")
                        {
                            shape.IncrementTop(-shift);

                            //shape. = "Оцениваемая квартира №"+appartmentNum.Text;
                        }
                        if (l == "podezd")
                        {
                            shape.IncrementTop(-shift);
                        }
                        if (l == "stairway")
                        {
                            shape.IncrementTop(-shift);
                        }
                    }*/

                    //

                    wdApp.ActiveDocument.SaveAs(saveFileDialog1.FileName);

                    wdDoc.Close();
                    wdApp.Documents.Close();
                    wdApp.Quit();
                }
            }
            catch (Exception except)
            {
                //MessageBox.Show(except.Message);
                saveState();

                //wdApp.Documents.Close();
                wdApp.Quit();
            }
        }

        private void saveAddsAppartaments_Click(object sender, EventArgs e)
        {
            string fileName;
            if (bankName.Text == "брр")
            {
                fileName = "приложение отчет номер" + contractNum.Text + "квартира " + appartmentNum.Text + " " +
                           street.Text + " " + houseNum.Text + " для " + bankName.Text;
                saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();
            }
            else
            {
                fileName = "приложение отчет " + contractNum.Text + "квартира " + appartmentNum.Text + " " + street.Text +
                           " " + houseNum.Text + " для " + bankName.Text;
                saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();
            }

            if (DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                wdApp = new Application();
                var wdDoc = new Document();
                string template = "\\шаблоны\\Приложение.doc";
                if (bankName.Text == "втб 24")
                {
                    template = "\\шаблоны\\ПриложениеВТБ24.doc";
                }
                wdDoc = wdApp.Documents.Open(System.Windows.Forms.Application.StartupPath + template);

                //button4.Text = System.Windows.Forms.Application.StartupPath + "\\template.doc";

                object replaceAll = WdReplace.wdReplaceAll;

                // Gets a NumberFormatInfo associated with the en-US culture.
                NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
                nfi.NumberDecimalDigits = 0;
                nfi.NumberGroupSeparator = " ";

                nfi.PositiveSign = "";

                string ownerFullName = ownerSurname.Text + " " + ownerName.Text + " " + ownerInit.Text;
                string customerFullName = customerSurname.Text + " " + customerName.Text + " " + customerInit.Text;
                string calculationDateStr = calculationDate.Text;
                int sentencesCount = wdDoc.Sentences.Count;

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@MO@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = MO.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@2.1.2.2@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[19].Cells[1].Value.ToString();
                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@houseType@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                if (houseType.Text == "Панельный")
                {
                    wdApp.Selection.Find.Replacement.Text = "жб плиты";
                }
                else
                {
                    wdApp.Selection.Find.Replacement.Text = houseType.Text.ToLower();
                }
                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@lm2@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = lm2text.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@m2@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = m2text.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@calculationDate@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = calculationDateStr;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerFullname@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerFullName;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*      if (newSentence.Contains("@@customerFullname@@"))
                      {
                          newSentence = newSentence.Replace("@@customerFullname@@", customerFullName);
                          changed = true;
                      }
               */
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerFullname@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerFullName;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@rooms@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = roomsAsString();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@appartmentNum@@"))
                {
                    newSentence = newSentence.Replace("@@appartmentNum@@", "№" + appartmentNum.Text);
                    changed = true;
                }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@appartmentNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = appartmentNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /* if (newSentence.Contains("@@town@@"))
                 {
                     newSentence = newSentence.Replace("@@town@@", town.Text);
                     changed = true;
                 }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@street@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = street.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@houseNum@@"))
                {
                    newSentence = newSentence.Replace("@@houseNum@@", houseNum.Text);
                    changed = true;
                }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@houseNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = houseNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@buildingNum@@"))
                {
                }*/
                string buildNum = null;
                if (buildingNum.Text != "")
                {
                    buildNum = "корп." + buildingNum.Text + ".";

                    //newSentence = newSentence.Replace("@@buildingNum@@", houseNum.Text);
                    //changed = true;
                }
                else
                {
                    buildNum = buildingNum.Text;

                    //changed = true;
                }

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@buildingNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = buildNum;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /* if (newSentence.Contains("@@customerAddress@@"))
                 {
                     newSentence = newSentence.Replace("@@customerAddress@@", customerAddres.Text);
                     changed = true;
                 }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerAddress@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerAddres.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@floor@@"))
                {
                    newSentence = newSentence.Replace("@@floor@@", floor.Value.ToString());
                    changed = true;
                }*/

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@floor@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = floor.Value.ToString().ToLower();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@floors@@"))
                {
                    newSentence = newSentence.Replace("@@floors@@", floors.Text);
                    changed = true;
                }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@floors@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = floors.Text.ToLower();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@town@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = town.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@cost@@"))
                {
                    newSentence = newSentence.Replace("@@cost@@", finalCostRounded.ToString());
                    changed = true;
                }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@cost@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = finalCostRounded.ToString("N", nfi);

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@contractNum@@"))
                {
                    newSentence = newSentence.Replace("@@contractNum@@", contractNum.Text);
                    changed = true;
                }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@contractNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = contractNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@contractDate@@"))
                {
                    newSentence = newSentence.Replace("@@contractDate@@", contractDate.Text);
                    changed = true;
                }
              */
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@contractDate@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = contractDate.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@customerName@@"))
                {
                    newSentence = newSentence.Replace("@@customerName@@", customerName.Text);
                    changed = true;
                }
                             */
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerName@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerName.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@customerInit@@"))
                {
                    newSentence = newSentence.Replace("@@customerInit@@", customerInit.Text);
                    changed = true;
                }
                 */
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerInit@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerInit.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                /* if (newSentence.Contains("@@likvidCost@@"))
        {
            newSentence = newSentence.Replace("@@likvidCost@@", likvidCost.ToString());
            changed = true;
        }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@likvidCost@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = likvidCost.ToString("N", nfi);

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                /* if (newSentence.Contains("@@stringCost@@"))
       {
           newSentence = newSentence.Replace("@@stringCost@@", costStr);
           changed = true;
       }
                */
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@stringCost@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = costStr;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*
if (newSentence.Contains("@@uvaj@@"))
{
    newSentence = newSentence.Replace("@@uvaj@@", uvaj);
    changed = true;
}*/
                getUvaj();
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@uvaj@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = uvaj;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (changed)
                {
                    wdDoc.Sentences[i].Text = newSentence;
                }
                                 */
                /*
                }

                /*int shapesCount = wdDoc.Shapes.Count;
                for (int i = 1; i <= shapesCount; i++)
                {
                //if (wdDoc.Shapes[i].
                if (wdDoc.Shapes[i].TextEffect.Text !=null)
                {
                    if (wdDoc.Shapes[i].TextEffect.Text.Contains("@@contractDate@@"))
                    {
                        wdDoc.Shapes[i].TextEffect.Text.Replace("@@contractDate@@", contractDate.Text);
                    }
                }
                }*/

                //Customer Passport
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerPassport@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerPassport.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerPassNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerPassNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerPassOVD@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerPassOVD.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerPassDate@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerPassDate.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerFullAddress@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerAddres.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                //owner Passport
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@passportSerial@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerPassport.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerPassNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerPassNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerPassOVD@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerPassOVD.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerPassDate@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerPassDate.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerFullAddress@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerAddress.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerDoc@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerDocs.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@registrationDoc@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = registrationDoc.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@tehPass@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[41].Cells[1].Value.ToString();
                ;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@b4.4@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = calculationAppartaments.Rows[38].Cells[2].Value.ToString();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@docType@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = docType.ToLower();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@2.1.2.20@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[37].Cells[1].Value.ToString();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@2.1.3.15@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = objectDataGrid.Rows[55].Cells[1].Value.ToString().ToLower();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                string te = wdApp.Selection.Text;

                //saving

                wdApp.ActiveDocument.SaveAs(saveFileDialog1.FileName);

                wdApp.Quit();
            }
        }

        private void saveAddsHouse_Click(object sender, EventArgs e)
        {
            string fileName;
            if (bankName.Text == "брр")
            {
                fileName = "приложение отчет номер" + contractNum.Text + "домовладение " + appartmentNum.Text + " " +
                           street.Text + " " + houseNum.Text + " для " + bankName.Text;
                saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();
            }
            else
            {
                fileName = "приложение отчет " + contractNum.Text + "домовладение и земельный участок " + appartmentNum.Text + " " +
                           street.Text + " " + houseNum.Text + " для " + bankName.Text;
                saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
                saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();
            }

            if (DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                wdApp = new Application();
                var wdDoc = new Document();

                wdDoc =
                    wdApp.Documents.Open(System.Windows.Forms.Application.StartupPath + "\\шаблоны\\ДомПриложение.doc");

                //button4.Text = System.Windows.Forms.Application.StartupPath + "\\template.doc";

                object replaceAll = WdReplace.wdReplaceAll;

                // Gets a NumberFormatInfo associated with the en-US culture.
                NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
                nfi.NumberDecimalDigits = 0;
                nfi.NumberGroupSeparator = " ";

                nfi.PositiveSign = "";

                string ownerFullName = ownerSurname.Text + " " + ownerName.Text + " " + ownerInit.Text;
                string customerFullName = customerSurname.Text + " " + customerName.Text + " " + customerInit.Text;
                string calculationDateStr = calculationDate.Text;
                int sentencesCount = wdDoc.Sentences.Count;

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@MO@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = MO.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@houseType@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                if (houseType.Text == "Панельный")
                {
                    wdApp.Selection.Find.Replacement.Text = "жб плиты";
                }
                else
                {
                    wdApp.Selection.Find.Replacement.Text = houseType.Text.ToLower();
                }
                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@lm2@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = lm2text.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@m2@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = m2text.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                ReplaceTextWord(ref wdApp, "@@dirtm2@@", dirtm2.Text);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@calculationDate@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = calculationDateStr;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerFullname@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerFullName;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*      if (newSentence.Contains("@@customerFullname@@"))
                      {
                          newSentence = newSentence.Replace("@@customerFullname@@", customerFullName);
                          changed = true;
                      }
               */
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerFullname@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerFullName;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@rooms@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = roomsAsString();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@appartmentNum@@"))
                {
                    newSentence = newSentence.Replace("@@appartmentNum@@", "№" + appartmentNum.Text);
                    changed = true;
                }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@appartmentNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = appartmentNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /* if (newSentence.Contains("@@town@@"))
                 {
                     newSentence = newSentence.Replace("@@town@@", town.Text);
                     changed = true;
                 }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@street@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = street.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@houseNum@@"))
                {
                    newSentence = newSentence.Replace("@@houseNum@@", houseNum.Text);
                    changed = true;
                }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@houseNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = houseNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@buildingNum@@"))
                {
                }*/
                string buildNum = null;
                if (buildingNum.Text != "")
                {
                    buildNum = "корп." + buildingNum.Text + ".";

                    //newSentence = newSentence.Replace("@@buildingNum@@", houseNum.Text);
                    //changed = true;
                }
                else
                {
                    buildNum = buildingNum.Text;

                    //changed = true;
                }

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@buildingNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = buildNum;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /* if (newSentence.Contains("@@customerAddress@@"))
                 {
                     newSentence = newSentence.Replace("@@customerAddress@@", customerAddres.Text);
                     changed = true;
                 }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerAddress@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerAddres.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@floor@@"))
                {
                    newSentence = newSentence.Replace("@@floor@@", floor.Value.ToString());
                    changed = true;
                }*/

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@floor@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = floor.Value.ToString().ToLower();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@floors@@"))
                {
                    newSentence = newSentence.Replace("@@floors@@", floors.Text);
                    changed = true;
                }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@floors@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = floors.Text.ToLower();

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@town@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = town.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@cost@@"))
                {
                    newSentence = newSentence.Replace("@@cost@@", finalCostRounded.ToString());
                    changed = true;
                }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@cost@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = finalCostRounded.ToString("N", nfi);

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                ReplaceTextWord(ref wdApp, "@@dirtCostR@@", dirtCalcGrid.Rows[32].Cells[1].Value.ToString());
                ReplaceTextWord(ref wdApp, "@@likvidCostDirt@@", dirtCalcGrid.Rows[33].Cells[1].Value.ToString());
                ReplaceTextWord(ref wdApp, "@@likvidCostFulle@@", (likvidCostDirt + likvidCost).ToString());
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@contractNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = contractNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@contractDate@@"))
                {
                    newSentence = newSentence.Replace("@@contractDate@@", contractDate.Text);
                    changed = true;
                }
              */
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@contractDate@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = contractDate.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@customerName@@"))
                {
                    newSentence = newSentence.Replace("@@customerName@@", customerName.Text);
                    changed = true;
                }
                             */
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerName@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerName.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (newSentence.Contains("@@customerInit@@"))
                {
                    newSentence = newSentence.Replace("@@customerInit@@", customerInit.Text);
                    changed = true;
                }
                 */
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerInit@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerInit.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                /* if (newSentence.Contains("@@likvidCost@@"))
        {
            newSentence = newSentence.Replace("@@likvidCost@@", likvidCost.ToString());
            changed = true;
        }*/
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@likvidCost@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = likvidCost.ToString("N", nfi);

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                //ReplaceTextWord(ref wdApp, "@@likvidCostDirt@@",

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@stringCost@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = costStr;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*
if (newSentence.Contains("@@uvaj@@"))
{
    newSentence = newSentence.Replace("@@uvaj@@", uvaj);
    changed = true;
}*/
                getUvaj();
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@uvaj@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = uvaj;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                /*if (changed)
                {
                    wdDoc.Sentences[i].Text = newSentence;
                }
                                 */
                /*
                }

                /*int shapesCount = wdDoc.Shapes.Count;
                for (int i = 1; i <= shapesCount; i++)
                {
                //if (wdDoc.Shapes[i].
                if (wdDoc.Shapes[i].TextEffect.Text !=null)
                {
                    if (wdDoc.Shapes[i].TextEffect.Text.Contains("@@contractDate@@"))
                    {
                        wdDoc.Shapes[i].TextEffect.Text.Replace("@@contractDate@@", contractDate.Text);
                    }
                }
                }*/

                //Customer Passport
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerPassport@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerPassport.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerPassNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerPassNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerPassOVD@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerPassOVD.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerPassDate@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerPassDate.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@customerFullAddress@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = customerAddres.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                //owner Passport
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@passportSerial@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerPassport.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerPassNum@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerPassNum.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerPassOVD@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerPassOVD.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerPassDate@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerPassDate.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "@@ownerFullAddress@@";
                wdApp.Selection.Find.Replacement.ClearFormatting();
                wdApp.Selection.Find.Replacement.Text = ownerAddress.Text;

                wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                //wdApp.Selection.Find.ClearFormatting();
                //wdApp.Selection.Find.Text = "@@ownerDoc@@";
                //wdApp.Selection.Find.Replacement.ClearFormatting();
                //wdApp.Selection.Find.Replacement.Text = ownerDocs.Text;

                //wdApp.Selection.Find.Execute(
                //    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                //    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                //    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                ReplaceTextWord(ref wdApp, "@@ownerDoc@@", ownerDocs.Text);

                //wdApp.Selection.Find.ClearFormatting();
                //wdApp.Selection.Find.Text = "@@registrationDoc@@";
                //wdApp.Selection.Find.Replacement.ClearFormatting();
                //wdApp.Selection.Find.Replacement.Text = registrationDoc.Text;

                //wdApp.Selection.Find.Execute(
                //    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                //    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                //    ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                ReplaceTextWord(ref wdApp, "@@registrationDoc@@", registrationDoc.Text);

                string te = wdApp.Selection.Text;

                //saving

                wdApp.ActiveDocument.SaveAs(saveFileDialog1.FileName);

                wdApp.Quit();
            }
        }

        private void roomsNum_ValueChanged(object sender, EventArgs e)
        {
            switch (docTypeT)
            {
                case "Квартира":
                    {
                        roomsAsString();
                        analogsGrid.Rows[0].Cells[1].Value = roomsX;
                        analogsGrid.Rows[0].Cells[2].Value = roomsX;
                        analogsGrid.Rows[0].Cells[3].Value = roomsX;
                        analogsGrid.Rows[0].Cells[4].Value = roomsX;
                        objectDataGrid.Rows[1].Cells[1].Value = fullAddress();
                    }
                    break;

                default:
                    break;
            }
        }

        private void analogsGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //copy costs to calc table
            try
            {
                string t = analogsGrid.Rows[15].Cells[2].Value.ToString();
                t = t.Replace(" ", "");
                calculationAppartaments.Rows[0].Cells[2].Value = t;
                t = analogsGrid.Rows[15].Cells[3].Value.ToString();
                t = t.Replace(" ", "");
                calculationAppartaments.Rows[0].Cells[3].Value = t;
                t = analogsGrid.Rows[15].Cells[4].Value.ToString();
                t = t.Replace(" ", "");
                calculationAppartaments.Rows[0].Cells[4].Value = t;

                //копирует площади в таблицу расчета
                t = analogsGrid.Rows[7].Cells[2].Value.ToString();
                t = t.Replace(" ", "");
                calculationAppartaments.Rows[1].Cells[2].Value = t;
                t = analogsGrid.Rows[7].Cells[3].Value.ToString();
                t = t.Replace(" ", "");
                calculationAppartaments.Rows[1].Cells[3].Value = t;
                t = analogsGrid.Rows[7].Cells[4].Value.ToString();
                t = t.Replace(" ", "");
                calculationAppartaments.Rows[1].Cells[4].Value = t;
            }
            catch (Exception exp)
            {
            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            new mapForm();
        }

        private void street_KeyUp(object sender, KeyEventArgs e)
        {
            if (docTypeT == "Квартира")
            {
                analogsGrid.Rows[2].Cells[1].Value = street.Text;
                analogsGrid.Rows[2].Cells[2].Value = street.Text;
                analogsGrid.Rows[2].Cells[3].Value = street.Text;
                analogsGrid.Rows[2].Cells[4].Value = street.Text;
            }
        }

        /// <summary>
        ///     Create xml file
        /// </summary>
        public void saveXML(string fileName)
        {
            try
            {
                var f = new FileStream(fileName, FileMode.OpenOrCreate);
                var settings = new XmlTextWriter(f, Encoding.Default);
                settings.WriteStartDocument();
                settings.WriteStartElement("test");
                addAtributeToXml(settings, customerName.Name, customerName.Text);
                addAtributeToXml(settings, customerInit.Name, customerInit.Text);
                addAtributeToXml(settings, customerAddres.Name, customerAddres.Text);
                addAtributeToXml(settings, customerPassDate.Name, customerPassDate.Text);
                addAtributeToXml(settings, customerPassNum.Name, customerPassNum.Text);
                addAtributeToXml(settings, customerPassOVD.Name, customerPassOVD.Text);
                addAtributeToXml(settings, customerPassport.Name, customerPassport.Text);
                addAtributeToXml(settings, customerPhone.Name, customerPhone.Text);
                addAtributeToXml(settings, customerSurname.Name, customerSurname.Text);
                addAtributeToXml(settings, ownerName.Name, ownerName.Text);
                addAtributeToXml(settings, ownerAddress.Name, ownerAddress.Text);
                addAtributeToXml(settings, ownerInit.Name, ownerInit.Text);
                addAtributeToXml(settings, ownerPassDate.Name, ownerPassDate.Text);
                addAtributeToXml(settings, ownerPassNum.Name, ownerPassNum.Text);
                addAtributeToXml(settings, ownerPassOVD.Name, ownerPassOVD.Text);
                addAtributeToXml(settings, ownerPassport.Name, ownerPassport.Text);
                addAtributeToXml(settings, ownerPhone.Name, ownerPhone.Text);
                addAtributeToXml(settings, ownerSurname.Name, ownerSurname.Text);
                addAtributeToXml(settings, town.Name, town.Text);
                addAtributeToXml(settings, street.Name, street.Text);
                addAtributeToXml(settings, buildingNum.Name, buildingNum.Text);
                addAtributeToXml(settings, roomsNum.Name, roomsNum.Text);
                addAtributeToXml(settings, appartmentNum.Name, appartmentNum.Text);
                addAtributeToXml(settings, calculationDate.Name, calculationDate.Text);
                addAtributeToXml(settings, contractDate.Name, contractDate.Text);
                addAtributeToXml(settings, contractNum.Name, contractNum.Text);
                addAtributeToXml(settings, floors.Name, floors.Text);
                addAtributeToXml(settings, floor.Name, floor.Text);
                addAtributeToXml(settings, houseType.Name, houseType.Text);
                addAtributeToXml(settings, houseNum.Name, houseNum.Text);
                addAtributeToXml(settings, registrationDoc.Name, registrationDoc.Text);
                addAtributeToXml(settings, MO.Name, MO.Text);
                addAtributeToXml(settings, ownerDocs.Name, ownerDocs.Text);
                addAtributeToXml(settings, m2text.Name, m2text.Text);
                addAtributeToXml(settings, lm2text.Name, lm2text.Text);
                if (docTypeT == "Квартира")
                {
                    addAtributeToXml(settings, "data2.1.1.2", objectDataGrid.Rows[2].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.3", objectDataGrid.Rows[3].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.4", objectDataGrid.Rows[4].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.5", objectDataGrid.Rows[5].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.6", objectDataGrid.Rows[6].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.7", objectDataGrid.Rows[7].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.8", objectDataGrid.Rows[8].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.9", objectDataGrid.Rows[9].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.10", objectDataGrid.Rows[10].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.11", objectDataGrid.Rows[11].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.12", objectDataGrid.Rows[12].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.13", objectDataGrid.Rows[13].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.14", objectDataGrid.Rows[14].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.15", objectDataGrid.Rows[15].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.1.16", objectDataGrid.Rows[16].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.1", objectDataGrid.Rows[18].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.2", objectDataGrid.Rows[19].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.3", objectDataGrid.Rows[20].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.4", objectDataGrid.Rows[21].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.5", objectDataGrid.Rows[22].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.6", objectDataGrid.Rows[23].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.7", objectDataGrid.Rows[24].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.8", objectDataGrid.Rows[25].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.9", objectDataGrid.Rows[26].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.10", objectDataGrid.Rows[27].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.11", objectDataGrid.Rows[28].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.12", objectDataGrid.Rows[29].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.13", objectDataGrid.Rows[30].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.14", objectDataGrid.Rows[31].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.15", objectDataGrid.Rows[32].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.16", objectDataGrid.Rows[33].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.17", objectDataGrid.Rows[34].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.18", objectDataGrid.Rows[35].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.19", objectDataGrid.Rows[36].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.20", objectDataGrid.Rows[37].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.21", objectDataGrid.Rows[38].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.2.22", objectDataGrid.Rows[39].Cells[1].Value.ToString());

                    addAtributeToXml(settings, "data2.1.3.1", objectDataGrid.Rows[41].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.2", objectDataGrid.Rows[42].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.3", objectDataGrid.Rows[43].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.4", objectDataGrid.Rows[44].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.5", objectDataGrid.Rows[45].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.6", objectDataGrid.Rows[46].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.7", objectDataGrid.Rows[47].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.8", objectDataGrid.Rows[48].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.9", objectDataGrid.Rows[49].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.10", objectDataGrid.Rows[50].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.11", objectDataGrid.Rows[51].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.12", objectDataGrid.Rows[52].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.13", objectDataGrid.Rows[53].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.14", objectDataGrid.Rows[54].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.15", objectDataGrid.Rows[55].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.16", objectDataGrid.Rows[56].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.17", objectDataGrid.Rows[57].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.18", objectDataGrid.Rows[58].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.19", objectDataGrid.Rows[59].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.20", objectDataGrid.Rows[60].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.21", objectDataGrid.Rows[61].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.22", objectDataGrid.Rows[62].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.23", objectDataGrid.Rows[63].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.24", objectDataGrid.Rows[64].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.25", objectDataGrid.Rows[65].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.26", objectDataGrid.Rows[66].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.27", objectDataGrid.Rows[67].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.28", objectDataGrid.Rows[68].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.29", objectDataGrid.Rows[69].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.30", objectDataGrid.Rows[70].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.31", objectDataGrid.Rows[71].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.32", objectDataGrid.Rows[72].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.33", objectDataGrid.Rows[73].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.34", objectDataGrid.Rows[74].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.35", objectDataGrid.Rows[75].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.36", objectDataGrid.Rows[76].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.37", objectDataGrid.Rows[77].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.38", objectDataGrid.Rows[78].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "data2.1.3.39", objectDataGrid.Rows[79].Cells[1].Value.ToString());
                    addAtributeToXml(settings, "analogsColsCount", analogsGrid.Columns.Count.ToString());
                    for (int i = 0; i < analogsGrid.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < analogsGrid.Columns.Count; j++)
                        {
                            addAtributeToXml(settings, "analog" + i.ToString() + "." + j.ToString(),
                                             analogsGrid.Rows[i].Cells[j].Value.ToString());
                        }
                    }

                    addAtributeToXml(settings, "calcColsCount", calculationAppartaments.Columns.Count.ToString());
                    for (int i = 0; i < calculationAppartaments.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < calculationAppartaments.Columns.Count; j++)
                        {
                            addAtributeToXml(settings, "calc" + i.ToString() + "." + j.ToString(),
                                             calculationAppartaments.Rows[i].Cells[j].Value.ToString());
                        }
                    }
                }

                settings.WriteEndElement();
                settings.WriteEndDocument();
                settings.Close();
                f.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);

                //this.Close();
            }
        }

        private void loadDataButton(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            loadState(openFileDialog1.FileName);
        }

        /// <summary>
        ///     подсчет кол-ва коэффициентов
        /// </summary>
        public int setCoefsCount(string coef, int counter)
        {
            if (coef != "")
            {
                double cellValue = double.Parse(coef);
                if (cellValue != 1.00)
                {
                    counter++;
                }
            }
            return counter;
        }

        private void objectDataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            analogsGrid.Rows[9].Cells[1].Value = objectDataGrid.Rows[53].Cells[1].Value;
            analogsGrid.Rows[9].Cells[2].Value = objectDataGrid.Rows[53].Cells[1].Value;
            analogsGrid.Rows[9].Cells[3].Value = objectDataGrid.Rows[53].Cells[1].Value;
            analogsGrid.Rows[9].Cells[4].Value = objectDataGrid.Rows[53].Cells[1].Value;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            analogsGrid.Rows[2].Cells[1].Value = street.Text + " " + textBox1.Text;
        }

        /// <summary>
        ///     Replace replaceText by text in word document
        /// </summary>
        public bool ReplaceTextWord(ref Application wdApp, string replaceText, string text)
        {
            try
            {
                if (text == null)
                {
                    text = "";
                }

                object replaceAll = WdReplace.wdReplaceAll;
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = replaceText;
                while (wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
                {
                    //wdNew = wdApp.Documents.Open(System.Windows.Forms.Application.StartupPath + "новостройка.doc", Missing, true);

                    //wdApp.Selection.Text = "";
                    wdApp.Selection.Text = text;

                    wdApp.ActiveDocument.Sections[1].Range.Select();
                }
                /*wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = replaceText;
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = text;

                    bool result = wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);*/
                /*if (!result)
                    {
                        return result;
                    }
                    else
                    {
                        throw new Exception("1");

                        //return false;
                    }*/
                return true;
            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
        }

        private void addOwner_Click(object sender, EventArgs e)
        {
            var currentOwner = new Owner(ownerAddress.Text,
                                         ownerPassOVD.Text, ownerInit.Text,
                                         ownerName.Text, ownerSurname.Text,
                                         ownerPassDate.Text,
                                         ownerPassNum.Text,
                                         ownerPassport.Text, ownerPhone.Text);
            owners.Add(currentOwner);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                /* string dbname = "", server = "", dbuser = "", dbpass = "";
                FileStream f = new FileStream("properties.xml", FileMode.OpenOrCreate);

                XmlTextReader settings = new XmlTextReader(f);
                while (settings.Read())
                {
                    if (settings.NodeType == XmlNodeType.Element)
                    {
                        if (settings.Name.Equals("server"))
                        {
                            server = settings.GetAttribute("servername");
                            dbname = settings.GetAttribute("dbname");
                            dbuser = settings.GetAttribute("dbuser");
                            dbpass = settings.GetAttribute("dbpass");
                        }
                    }
                }
                f.Close();

                string CommandText = "select accessLevel, idusers from ads_paper.users";
                string Connect = "Database=" + dbname + ";Data Source=" + server + ";User Id=" + dbuser + ";Password=" + dbpass;

                //Переменная Connect - это строка подключения в которой:
                //БАЗА - Имя базы в MySQL
                //ХОСТ - Имя или IP-адрес сервера (если локально то можно и localhost)
                //ПОЛЬЗОВАТЕЛЬ - Имя пользователя MySQL
                //ПАРОЛЬ - говорит само за себя - пароль пользователя БД MySQL

                MySqlConnection myConnection = new MySqlConnection(Connect);
                MySqlCommand myCommand = new MySqlCommand(CommandText, myConnection);
                myConnection.Open(); //Устанавливаем соединение с базой данных.
                MySqlDataReader MyDataReader;

                MyDataReader = myCommand.ExecuteReader();

               if (MyDataReader.Read())
                {
                    if (!MyDataReader.IsDBNull(0))
                    {
                       // MessageBox.Show("Подключение к базе данных прошло успешно");
                    }
                    else
                    {
                      //  MessageBox.Show("Не удалось подключиться к базе данных");
                    }
                }
                MyDataReader.Close();
                myConnection.Close(); //Обязательно закрываем соединение!
                * */
                saveState();
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message);
            }
        }

        private void showOwnersList_Click(object sender, EventArgs e)
        {
            new ownerForm(owners);
        }

        /// <summary>
        ///     Вставка переводов строк вместо \rn
        /// </summary>
        public bool InsertParagraphs(ref Application wdApp)
        {
            try
            {
                wdApp.Selection.Find.ClearFormatting();
                wdApp.Selection.Find.Text = "/rn";

                while (wdApp.Selection.Find.Execute(
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
                {
                    //Microsoft.Office.Interop.Word.Range r1;

                    //r1.Text = objectDataGrid.Rows[63].Cells[1].Value.ToString();
                    wdApp.Selection.InsertParagraph();

                    //wdApp.Selection.Font.Superscript = 1;
                    //                        wdApp.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault);
                    wdApp.ActiveDocument.Sections[1].Range.Select();
                }
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();

            //Microsoft.Office.Interop.Excel.Workbook excelDoc = new Microsoft.Office.Interop.Excel.Workbook();
            string ownerFullName = ownerSurname.Text + " " + ownerName.Text + " " + ownerInit.Text;
            string customerFullName = customerSurname.Text + " " + customerName.Text + " " + customerInit.Text;
            string fileName = "отчет " + contractNum.Text + " расчет стоимости домовладения " + street.Text + " " +
                              houseNum.Text + " для " + bankName.Text;
            saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();

            if (DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                excelApp.Workbooks.Open(System.Windows.Forms.Application.StartupPath + "\\дом.xls", Missing, Missing,
                                        Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
                                        Missing, Missing, Missing);
                int analogsCount = houseCalcGrid.ColumnCount;
                int rowCount = houseCalcGrid.RowCount;
                for (int j = 1; j < analogsCount; j++)
                {
                    for (int i = 1; i < rowCount; i++)
                    {
                        if (houseCalcGrid.Rows[i - 1].Cells[j].Value != null)
                        {
                            excelApp.Workbooks[1].Sheets[1].Cells[i + 1, j + 1] =
                                houseCalcGrid.Rows[i - 1].Cells[j].Value.ToString();
                        }
                    }
                }

                //первый аналог

                //string pattern = "MMMM yyyyг.";
                //string d1 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[2].Value.ToString()).ToString(pattern);
                //string d2 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[3].Value.ToString()).ToString(pattern);
                //string d3 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[4].Value.ToString()).ToString(pattern);
                //excelApp.Workbooks[1].Sheets[1].Cells[7, 3] = d1;
                //excelApp.Workbooks[1].Sheets[1].Cells[7, 4] = d2;
                //excelApp.Workbooks[1].Sheets[1].Cells[7, 5] = d3;

                excelApp.ActiveWorkbook.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing,
                                               XlSaveAsAccessMode.xlNoChange,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelApp.ActiveWorkbook.Close();
                excelApp.Quit();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();

            //Microsoft.Office.Interop.Excel.Workbook excelDoc = new Microsoft.Office.Interop.Excel.Workbook();
            string ownerFullName = ownerSurname.Text + " " + ownerName.Text + " " + ownerInit.Text;
            string customerFullName = customerSurname.Text + " " + customerName.Text + " " + customerInit.Text;
            string fileName = "отчет " + contractNum.Text + " расчет стоимости земельного участка " + street.Text + " " +
                              houseNum.Text + " для " + bankName.Text;
            saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();

            if (DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                excelApp.Workbooks.Open(System.Windows.Forms.Application.StartupPath + "\\земля.xls", Missing, Missing,
                                        Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
                                        Missing, Missing, Missing);
                int analogsCount = dirtCalcGrid.ColumnCount;
                int rowCount = dirtCalcGrid.RowCount;
                for (int j = 1; j < analogsCount; j++)
                {
                    for (int i = 1; i < rowCount; i++)
                    {
                        if (dirtCalcGrid.Rows[i - 1].Cells[j].Value != null)
                        {
                            excelApp.Workbooks[1].Sheets[1].Cells[i + 1, j + 1] =
                                dirtCalcGrid.Rows[i - 1].Cells[j].Value.ToString();
                        }
                    }
                }

                //первый аналог

                //string pattern = "MMMM yyyyг.";
                //string d1 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[2].Value.ToString()).ToString(pattern);
                //string d2 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[3].Value.ToString()).ToString(pattern);
                //string d2 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[3].Value.ToString()).ToString(pattern);
                //string d2 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[3].Value.ToString()).ToString(pattern);
                //string d2 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[3].Value.ToString()).ToString(pattern);
                //string d2 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[3].Value.ToString()).ToString(pattern);
                //string d3 = Convert.ToDateTime(analogsGrid.Rows[18].Cells[4].Value.ToString()).ToString(pattern);
                //excelApp.Workbooks[1].Sheets[1].Cells[7, 3] = d1;
                //excelApp.Workbooks[1].Sheets[1].Cells[7, 4] = d2;
                //excelApp.Workbooks[1].Sheets[1].Cells[7, 5] = d3;

                excelApp.ActiveWorkbook.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing,
                                               XlSaveAsAccessMode.xlNoChange,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelApp.ActiveWorkbook.Close();
                excelApp.Quit();
            }
        }

        private void dirtGridAnalogs_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            //площади
            dirtCalcGrid.Rows[2].Cells[1].Value = dirtGridAnalogs.Rows[1].Cells[1].Value;
            dirtCalcGrid.Rows[2].Cells[2].Value = dirtGridAnalogs.Rows[1].Cells[2].Value;
            dirtCalcGrid.Rows[2].Cells[3].Value = dirtGridAnalogs.Rows[1].Cells[3].Value;

            //цена
            dirtCalcGrid.Rows[1].Cells[1].Value = dirtGridAnalogs.Rows[2].Cells[1].Value;
            dirtCalcGrid.Rows[1].Cells[2].Value = dirtGridAnalogs.Rows[2].Cells[2].Value;
            dirtCalcGrid.Rows[1].Cells[3].Value = dirtGridAnalogs.Rows[2].Cells[3].Value;

            calculateCostDirt();
        }

        private void dirtm2_TextChanged(object sender, EventArgs e)
        {
            dirtCalcGrid.Rows[30].Cells[1].Value = dirtm2.Text;
        }

        private void saveGridToWordButton_Click(object sender, EventArgs e)
        {
            //HouseCostCalculation.House h = new HouseCostCalculation.House();
            //h.saveHouse(this);
            string townName = " " + town.Text + ", ";

            if ((town.Text == "г. Владикавказ") || (town.Text == "г.Владикавказ"))
            {
                townName = " ";
            }

            string buildNum = null;

            if (buildingNum.Text != "")
            {
                buildNum = "корп. " + buildingNum.Text;
            }
            roomsAsString();
            string fileName = "отчет номер " + contractNum.Text + " от " + calculationDate.Text + " договор от" +
                              contractDate.Text + " " + fullAddressDirt() + " " + ownerSurname.Text + " " +
                              ownerName.Text + " для " + customerSurname.Text + " " + customerName.Text + " " +
                              bankName.Text + ".doc";
            saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
            saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();
            try
            {
                if (DialogResult.OK == saveFileDialog1.ShowDialog())
                {
                    wdApp = new Application();
                    var wdDoc = new Document();

                    wdDoc = wdApp.Documents.Open(System.Windows.Forms.Application.StartupPath + "\\земля.doc", Missing,
                                                 true);
                    object replaceAll = WdReplace.wdReplaceAll;

                    // Gets a NumberFormatInfo associated with the en-US culture.
                    NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;

                    nfi.NumberDecimalDigits = 0;
                    nfi.NumberGroupSeparator = " ";

                    nfi.PositiveSign = "";

                    string ownerFullName = ownerSurname.Text + " " + ownerName.Text + " " + ownerInit.Text;
                    string customerFullName = customerSurname.Text + " " + customerName.Text + " " + customerInit.Text;

                    calculationDate.CustomFormat = "dd MMMM yyyy";
                    string calculationDateStr = calculationDate.Text;
                    int lenght = calculationDateStr.Length;
                    string temp = null;
                    string t;

                    for (int i = 0; i < lenght; i++)
                    {
                        if (i == 3)
                        {
                            t = calculationDateStr[i].ToString().ToUpper();
                            temp += t;
                        }
                        else
                        {
                            temp += calculationDateStr[i];
                        }
                    }

                    calculationDateStr = temp;

                    calculationDate.CustomFormat = "dd/MM/yy";
                    int sentencesCount = wdDoc.Sentences.Count;
                    string topColontitul = topColontitulCreatorHouse();

                    wdDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = topColontitul;

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@MO@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = MO.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@dirtCost@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dirtCalcGrid.Rows[31].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@cost@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dirtCalcGrid.Rows[31].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@dirtm2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dirtm2.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@dirtCostR@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dirtCalcGrid.Rows[32].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@calculationDateStr@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationDateStr;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@houseType@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = houseType.Text.ToLower();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@roomsT@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();

                    wdApp.Selection.Find.Replacement.Text = roomsT;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@roomsX@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = roomsX;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@lm2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = lm2text.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@m2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = m2text.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerNameInits@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerFamiliyR + " " + getInits();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@calculationDate@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = calculationDate.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@ownerFullname@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerFullName;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerFullname@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerFullName;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@rooms1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    roomsAsString();
                    wdApp.Selection.Find.Replacement.Text = rooms1;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@ownerFullnameR@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerFullNameR;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerFullnameR@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerFullNameR;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerFullnameT@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerFullNameT;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@ownerFullnameD@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerFullNameD;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@ownerFullnameT@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerFullNameT;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*
                    /*
                       if (newSentence.Contains("@@customerFullnameD@@"))
                       {
                           newSentence = newSentence.Replace("@@customerFullnameD@@", customerFullNameD);
                           changed = true;
                       }*/
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerFullnameD@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerFullNameD;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@rooms@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = roomsAsString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@appartmentNum@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = "№" + appartmentNum.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /* if (newSentence.Contains("@@town@@"))
                     {
                         newSentence = newSentence.Replace("@@town@@", town.Text);
                         changed = true;
                     }*/
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@street@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = street.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*if (newSentence.Contains("@@houseNum@@"))
                    {
                        newSentence = newSentence.Replace("@@houseNum@@", houseNum.Text);
                        changed = true;
                    }*/
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@houseNum@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = houseNum.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    buildNum = null;
                    if (buildingNum.Text != "")
                    {
                        buildNum = " корп." + buildingNum.Text;
                    }
                    else
                    {
                        buildNum = buildingNum.Text;
                    }

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@buildingNum@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = buildNum;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /* if (newSentence.Contains("@@customerAddress@@"))
                     {
                         newSentence = newSentence.Replace("@@customerAddress@@", customerAddres.Text);
                         changed = true;
                     }*/
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerAddress@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerAddres.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*if (newSentence.Contains("@@floor@@"))
                    {
                        newSentence = newSentence.Replace("@@floor@@", floor.Value.ToString());
                        changed = true;
                    }*/

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@floor@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = floor.Value.ToString();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*if (newSentence.Contains("@@floors@@"))
                    {
                        newSentence = newSentence.Replace("@@floors@@", floors.Text);
                        changed = true;
                    }*/
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@floors@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = floors.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@town@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = town.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@cost@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = finalCostRounded.ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*if (newSentence.Contains("@@contractNum@@"))
                    {
                        newSentence = newSentence.Replace("@@contractNum@@", contractNum.Text);
                        changed = true;
                    }*/
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@contractNum@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = contractNum.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*if (newSentence.Contains("@@contractDate@@"))
                    {
                        newSentence = newSentence.Replace("@@contractDate@@", contractDate.Text);
                        changed = true;
                    }
                  */
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@contractDate@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = contractDate.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*if (newSentence.Contains("@@customerName@@"))
                    {
                        newSentence = newSentence.Replace("@@customerName@@", customerName.Text);
                        changed = true;
                    }
                                 */
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerName@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerName.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*if (newSentence.Contains("@@customerInit@@"))
                    {
                        newSentence = newSentence.Replace("@@customerInit@@", customerInit.Text);
                        changed = true;
                    }
                     */
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerInit@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerInit.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    /* if (newSentence.Contains("@@likvidCost@@"))
            {
                newSentence = newSentence.Replace("@@likvidCost@@", likvidCost.ToString());
                changed = true;
            }*/
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@likvidCost@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = likvidCost.ToString("N", nfi);

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    /* if (newSentence.Contains("@@stringCost@@"))
           {
               newSentence = newSentence.Replace("@@stringCost@@", costStr);
               changed = true;
           }
                    */
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@stringCost@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = costStr.ToLower();

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*
    if (newSentence.Contains("@@uvaj@@"))
    {
        newSentence = newSentence.Replace("@@uvaj@@", uvaj);
        changed = true;
    }*/
                    getUvaj();
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@uvaj@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = uvaj;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*if (changed)
                    {
                        wdDoc.Sentences[i].Text = newSentence;
                    }
                                     */
                    /*
                    }

                    /*int shapesCount = wdDoc.Shapes.Count;
                    for (int i = 1; i <= shapesCount; i++)
                    {
                    //if (wdDoc.Shapes[i].
                    if (wdDoc.Shapes[i].TextEffect.Text !=null)
                    {
                        if (wdDoc.Shapes[i].TextEffect.Text.Contains("@@contractDate@@"))
                        {
                            wdDoc.Shapes[i].TextEffect.Text.Replace("@@contractDate@@", contractDate.Text);
                        }
                    }
                    }*/

                    //Customer Passport
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerPassport@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerPassport.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerPassNum@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerPassNum.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerPassOVD@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerPassOVD.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerPassDate@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerPassDate.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@customerFullAddress@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = customerAddres.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    //owner Passport
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@passportSerial@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerPassport.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@ownerPassNum@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerPassNum.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@ownerPassOVD@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerPassOVD.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@ownerPassDate@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerPassDate.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@ownerFullAddress@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerAddress.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@ownerDoc@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = ownerDocs.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@registrationDoc@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = registrationDoc.Text;

                    wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    /*
                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@tehPass@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[41].Cells[1].Value.ToString(); ;

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[2].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[3].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[4].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[5].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[6].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[7].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[8].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[9].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[10].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[11].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[12].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[13].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[14].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.1.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[15].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.1@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[17].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.2@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[18].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.3@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[19].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.4@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[20].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.5@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[21].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.6@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[22].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.7@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[23].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.8@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[24].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.9@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[25].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.10@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[26].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.11@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[27].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.12@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[28].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.13@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[29].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.14@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[30].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.15@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[31].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.16@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[32].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.17@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[33].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.18@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[34].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.19@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[35].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.20@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[36].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.21@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[37].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.22@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[38].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.23@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[39].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.24@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[40].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.25@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[41].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.26@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[42].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.27@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[43].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.28@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[44].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.29@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[45].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.30@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[46].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.31@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[47].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.32@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[48].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.33@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[49].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.34@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[50].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.35@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[51].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.36@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[52].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.37@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[53].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.38@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[54].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.39@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[55].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.40@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[56].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "@@2.1.2.41@@";
                    wdApp.Selection.Find.Replacement.ClearFormatting();
                    wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[57].Cells[1].Value.ToString();

                    wdApp.Selection.Find.Execute(
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
                    */

                    ReplaceTextWord(ref wdApp, "@@a1.1@@", dirtGridAnalogs.Rows[0].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.2@@", dirtGridAnalogs.Rows[1].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.3@@", dirtGridAnalogs.Rows[2].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.4@@", dirtGridAnalogs.Rows[3].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.5@@", dirtGridAnalogs.Rows[4].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.6@@", dirtGridAnalogs.Rows[5].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.7@@", dirtGridAnalogs.Rows[6].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.8@@", dirtGridAnalogs.Rows[7].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.9@@", dirtGridAnalogs.Rows[8].Cells[1].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a1.10@@", dirtGridAnalogs.Rows[9].Cells[1].Value.ToString());

                    ReplaceTextWord(ref wdApp, "@@a2.1@@", dirtGridAnalogs.Rows[0].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.2@@", dirtGridAnalogs.Rows[1].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.3@@", dirtGridAnalogs.Rows[2].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.4@@", dirtGridAnalogs.Rows[3].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.5@@", dirtGridAnalogs.Rows[4].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.6@@", dirtGridAnalogs.Rows[5].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.7@@", dirtGridAnalogs.Rows[6].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.8@@", dirtGridAnalogs.Rows[7].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.9@@", dirtGridAnalogs.Rows[8].Cells[2].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a2.10@@", dirtGridAnalogs.Rows[9].Cells[2].Value.ToString());

                    ReplaceTextWord(ref wdApp, "@@a3.1@@", dirtGridAnalogs.Rows[0].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.2@@", dirtGridAnalogs.Rows[1].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.3@@", dirtGridAnalogs.Rows[2].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.4@@", dirtGridAnalogs.Rows[3].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.5@@", dirtGridAnalogs.Rows[4].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.6@@", dirtGridAnalogs.Rows[5].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.7@@", dirtGridAnalogs.Rows[6].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.8@@", dirtGridAnalogs.Rows[7].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.9@@", dirtGridAnalogs.Rows[8].Cells[3].Value.ToString());
                    ReplaceTextWord(ref wdApp, "@@a3.10@@", dirtGridAnalogs.Rows[9].Cells[3].Value.ToString());

                    wdApp.Selection.Find.ClearFormatting();
                    wdApp.Selection.Find.Text = "м2";
                    while (wdApp.Selection.Find.Execute(
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
                        ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
                    {
                        wdApp.Selection.Characters[2].Font.Superscript = 1;
                    }
                    string te = wdApp.Selection.Text;

                    //saving
                    try
                    {
                        int x = wdDoc.Shapes.Count;
                        x = wdDoc.Shapes.Count;
                        for (int k = 1; k < x; k++)
                        {
                            Shape shape = wdDoc.Shapes[k];

                            //string l = shape.AlternativeText;
                            if (shape.AlternativeText.Contains("cont"))
                            {
                                wdDoc.Shapes[k].TextEffect.Text = "№ " + contractNum.Text + " от " +
                                                                  calculationDate.Text + "г.";
                            }
                        }
                    }
                    catch (Exception exp)
                    {
                    }

                    wdApp.ActiveDocument.SaveAs(saveFileDialog1.FileName);

                    wdApp.Quit();
                }
            }
            catch (Exception except)
            {
                wdApp.Quit();
                MessageBox.Show(except.Message);
            }
        }

        private void ownerDocs_TextChanged(object sender, EventArgs e)
        {
        }

        private void dirtCalcGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void houseCalcGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dirtGridAnalogs_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void houseCalcGrid_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dirtCalcGrid_Enter(object sender, EventArgs e)
        {
            try
            {
                if (dirtCalcGrid.SelectedCells.Count > 0)
                {
                    if (!dirtCalcGrid.SelectedCells[0].Value.ToString().Contains('.'))
                    {
                        calculateCostDirt();
                    }
                    else
                    {
                        MessageBox.Show("Была введена точка, вместо запятой");
                    }
                }
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message);
            }
        }

        private void houseCalcGrid_Enter(object sender, EventArgs e)
        {
            try
            {
                if (houseCalcGrid.SelectedCells.Count > 0)
                {
                    if (!houseCalcGrid.SelectedCells[0].Value.ToString().Contains('.'))
                    {
                        calculateCostHouse();
                    }
                    else
                    {
                        MessageBox.Show("Была введена точка, вместо запятой");
                    }
                }
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message);
            }
        }

        private void dirtCalcGrid_CellLeave_1(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dirtCalcGrid_CellEndEdit_1(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dirtCalcGrid.SelectedCells.Count > 0)
                {
                    if (!dirtCalcGrid.SelectedCells[0].Value.ToString().Contains('.'))
                    {
                        calculateCostDirt();
                    }
                    else
                    {
                        MessageBox.Show("Была введена точка, вместо запятой");
                    }
                }
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message);
            }
        }

        private void houseCalcGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (houseCalcGrid.SelectedCells.Count > 0)
                {
                    if (!houseCalcGrid.SelectedCells[0].Value.ToString().Contains('.'))
                    {
                        calculateCostHouse();
                    }
                    else
                    {
                        MessageBox.Show("Была введена точка, вместо запятой");
                    }
                }
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message);
            }
        }

        private void dirtm2_TextChanged_1(object sender, EventArgs e)
        {
            dirtCalcGrid.Rows[30].Cells[1].Value = dirtm2.Text;
            calculateCostDirt();
        }

        private void houseAnalogs_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //площади
            string t = houseAnalogs.Rows[6].Cells[2].Value.ToString();
            t = t.Replace(" ", "");
            houseCalcGrid.Rows[4].Cells[2].Value = t;
            t = houseAnalogs.Rows[6].Cells[3].Value.ToString();
            t = t.Replace(" ", "");
            houseCalcGrid.Rows[4].Cells[3].Value = t;
            t = houseAnalogs.Rows[6].Cells[4].Value.ToString();
            t = t.Replace(" ", "");
            houseCalcGrid.Rows[4].Cells[4].Value = t;

            //площадь земли
            t = houseAnalogs.Rows[9].Cells[2].Value.ToString();
            t = t.Replace(" ", "");
            houseCalcGrid.Rows[1].Cells[2].Value = t;
            t = houseAnalogs.Rows[9].Cells[3].Value.ToString();
            t = t.Replace(" ", "");
            houseCalcGrid.Rows[1].Cells[3].Value = t;
            t = houseAnalogs.Rows[9].Cells[4].Value.ToString();
            t = t.Replace(" ", "");
            houseCalcGrid.Rows[1].Cells[4].Value = t;

            //цена
            t = houseAnalogs.Rows[14].Cells[2].Value.ToString();
            t = t.Replace(" ", "");
            houseCalcGrid.Rows[0].Cells[2].Value = t;
            t = houseAnalogs.Rows[14].Cells[3].Value.ToString();
            t = t.Replace(" ", "");
            houseCalcGrid.Rows[0].Cells[3].Value = t;
            t = houseAnalogs.Rows[14].Cells[4].Value.ToString();
            t = t.Replace(" ", "");
            houseCalcGrid.Rows[0].Cells[4].Value = t;
            calculateCostHouse();
        }

        private void mainForm_Load(object sender, EventArgs e)
        {
            updateTables();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}