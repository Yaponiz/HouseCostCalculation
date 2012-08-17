using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HouseCostCalculation.Properties;
using WindowsFormsApplication1;
using HouseCostCalculation;

namespace HouseCostCalculation
{
    class House
    {
        public double calculateCost()
        {
            double cost = 0;
            return cost;
        }
            
        private Address adress;
        private double cost;
        
        public double Cost
        {
            get
            {
                return cost;
            }
            set
            {
                cost = value;
            }
        }
        public Address Address
        {
            get
            {
                return adress;
            }
            set
            {
                adress = value;
            }
        }

        public bool saveHouse(mainForm mainForm)
        {

            {
                
                //string townName = " " + mainForm.Controls["town"].Text + ", ";

                //if ((mainForm.Controls["town"].Text == "г. Владикавказ") || (mainForm.Controls["town"].Text == "г.Владикавказ"))
                //{
                //    townName = " ";
                //}

                //string buildNum = null;

                //if (mainForm.Controls["buildingNum"].Text != "")
                //{
                //    buildNum = "корп. " + mainForm.Controls["buildingNum"].Text;
                //}
                //mainForm.roomsAsString();
                //string fullAddress = mainForm.fullAddress();
                //string fileName = "отчет номер " + mainForm.Controls["contractNum"].Text + " от " + mainForm.Controls["calculationDate"].Text + " договор от" + mainForm.Controls["contractDate"].Text + " " + fullAddress + " " + mainForm.Controls["ownerSurname"].Text + " " + mainForm.Controls["ownerName"].Text + " для " + mainForm.Controls["customerSurname"].Text + " " + mainForm.Controls["customerName"].Text + " " + mainForm.Controls["bankName"].Text + ".doc";
               return true;
        //        mainForm.saveFileDialog1.FileName = fileName.Replace("\"", " ").ToLower();
        //        mainForm.saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("/", " ").ToLower();
        //        mainForm.saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(",", " ").ToLower();
        //        mainForm.saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("№", " ").ToLower();
        //        mainForm.saveFileDialog1.FileName = saveFileDialog1.FileName.Replace(".", " ").ToLower();
        //        saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("-", " ").ToLower();
        //        saveFileDialog1.FileName = saveFileDialog1.FileName.Replace("  ", " ").ToLower();
        //        try
        //        {
        //            if (DialogResult.OK == saveFileDialog1.ShowDialog())
        //            {
        //                wdApp = new Microsoft.Office.Interop.Word.Application();
        //                Microsoft.Office.Interop.Word.Document wdDoc = new Microsoft.Office.Interop.Word.Document();

        //                wdDoc = wdApp.Documents.Open(System.Windows.Forms.Application.StartupPath + "\\house.doc", Missing, true);
        //                object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

        //                // Gets a NumberFormatInfo associated with the en-US culture.
        //                NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;

        //                nfi.NumberDecimalDigits = 0;
        //                nfi.NumberGroupSeparator = " ";

        //                nfi.PositiveSign = "";



        //                string ownerFullName = ownerSurname.Text + " " + ownerName.Text + " " + ownerInit.Text;
        //                string customerFullName = customerSurname.Text + " " + customerName.Text + " " + customerInit.Text;

        //                calculationDate.CustomFormat = "dd MMMM yyyy";
        //                string calculationDateStr = calculationDate.Text;
        //                int lenght = calculationDateStr.Length;
        //                string temp = null;
        //                string t;

        //                for (int i = 0; i < lenght; i++)
        //                {
        //                    if (i == 3)
        //                    {
        //                        t = calculationDateStr[i].ToString().ToUpper();
        //                        temp += t;
        //                    }
        //                    else
        //                    {
        //                        temp += calculationDateStr[i];
        //                    }
        //                }

        //                calculationDateStr = temp;

        //                calculationDate.CustomFormat = "dd/MM/yy";
        //                int sentencesCount = wdDoc.Sentences.Count;
        //                string topColontitul = topColontitulCreator();


        //                wdDoc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = topColontitul;


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@MO@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = MO.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@dirtCost@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dirtCalcGrid.Rows[31].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@dirtCostR@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dirtCalcGrid.Rows[32].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@calculationDateStr@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = calculationDateStr;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@houseType@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = houseType.Text.ToLower();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@roomsT@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();


        //                wdApp.Selection.Find.Replacement.Text = roomsT;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@roomsX@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = roomsX;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@lm2@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = lm2text.Text;

        //                wdApp.Selection.Find.Execute(
        //             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@m2@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = m2text.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerNameInits@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerFamiliyR + " " + getInits();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@calculationDate@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = calculationDate.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@ownerFullname@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = ownerFullName;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerFullname@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerFullName;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@rooms1@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                roomsAsString();
        //                wdApp.Selection.Find.Replacement.Text = rooms1;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@ownerFullnameR@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = ownerFullNameR;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerFullnameR@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerFullNameR;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerFullnameT@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerFullNameT;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@ownerFullnameD@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = ownerFullNameD;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@ownerFullnameT@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = ownerFullNameT;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*
        //                /*
        //                   if (newSentence.Contains("@@customerFullnameD@@"))
        //                   {
        //                       newSentence = newSentence.Replace("@@customerFullnameD@@", customerFullNameD);
        //                       changed = true;
        //                   }*/
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerFullnameD@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerFullNameD;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@rooms@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = roomsAsString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@appartmentNum@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = "№" + appartmentNum.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /* if (newSentence.Contains("@@town@@"))
        //                 {
        //                     newSentence = newSentence.Replace("@@town@@", town.Text);
        //                     changed = true;
        //                 }*/
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@town@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = town.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*if (newSentence.Contains("@@houseNum@@"))
        //                {
        //                    newSentence = newSentence.Replace("@@houseNum@@", houseNum.Text);
        //                    changed = true;
        //                }*/
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@houseNum@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = houseNum.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                buildNum = null;
        //                if (buildingNum.Text != "")
        //                {
        //                    buildNum = " корп." + buildingNum.Text;

        //                }
        //                else
        //                {
        //                    buildNum = buildingNum.Text;
        //                }



        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@buildingNum@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = buildNum;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /* if (newSentence.Contains("@@customerAddress@@"))
        //                 {
        //                     newSentence = newSentence.Replace("@@customerAddress@@", customerAddres.Text);
        //                     changed = true;
        //                 }*/
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerAddress@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerAddres.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*if (newSentence.Contains("@@floor@@"))
        //                {
        //                    newSentence = newSentence.Replace("@@floor@@", floor.Value.ToString());
        //                    changed = true;
        //                }*/

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@floor@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = floor.Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*if (newSentence.Contains("@@floors@@"))
        //                {
        //                    newSentence = newSentence.Replace("@@floors@@", floors.Text);
        //                    changed = true;
        //                }*/
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@floors@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = floors.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@town@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = town.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@cost@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = finalCostRounded.ToString("N", nfi);

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*if (newSentence.Contains("@@contractNum@@"))
        //                {
        //                    newSentence = newSentence.Replace("@@contractNum@@", contractNum.Text);
        //                    changed = true;
        //                }*/
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@contractNum@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = contractNum.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*if (newSentence.Contains("@@contractDate@@"))
        //                {
        //                    newSentence = newSentence.Replace("@@contractDate@@", contractDate.Text);
        //                    changed = true;
        //                }
        //              */
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@contractDate@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = contractDate.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*if (newSentence.Contains("@@customerName@@"))
        //                {
        //                    newSentence = newSentence.Replace("@@customerName@@", customerName.Text);
        //                    changed = true;
        //                }
        //                             */
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerName@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerName.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*if (newSentence.Contains("@@customerInit@@"))
        //                {
        //                    newSentence = newSentence.Replace("@@customerInit@@", customerInit.Text);
        //                    changed = true;
        //                }
        //                 */
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerInit@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerInit.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                /* if (newSentence.Contains("@@likvidCost@@"))
        //        {
        //            newSentence = newSentence.Replace("@@likvidCost@@", likvidCost.ToString());
        //            changed = true;
        //        }*/
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@likvidCost@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = likvidCost.ToString("N", nfi);

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                /* if (newSentence.Contains("@@stringCost@@"))
        //       {
        //           newSentence = newSentence.Replace("@@stringCost@@", costStr);
        //           changed = true;
        //       }
        //                */
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@stringCost@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = costStr.ToLower();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*
        //if (newSentence.Contains("@@uvaj@@"))
        //{
        //    newSentence = newSentence.Replace("@@uvaj@@", uvaj);
        //    changed = true;
        //}*/
        //                getUvaj();
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@uvaj@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = uvaj;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                /*if (changed)
        //                {
        //                    wdDoc.Sentences[i].Text = newSentence;
        //                }
        //                                 */
        //                /*
        //                }

        //                /*int shapesCount = wdDoc.Shapes.Count;
        //                for (int i = 1; i <= shapesCount; i++)
        //                {
        //                //if (wdDoc.Shapes[i].
        //                if (wdDoc.Shapes[i].TextEffect.Text !=null)
        //                {
        //                    if (wdDoc.Shapes[i].TextEffect.Text.Contains("@@contractDate@@"))
        //                    {
        //                        wdDoc.Shapes[i].TextEffect.Text.Replace("@@contractDate@@", contractDate.Text);

        //                    }
        //                }

               
        //                }*/
        //                //Customer Passport
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerPassport@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerPassport.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerPassNum@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerPassNum.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerPassOVD@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerPassOVD.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerPassDate@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerPassDate.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@customerFullAddress@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = customerAddres.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //owner Passport
        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@passportSerial@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = passportSerial.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@ownerPassNum@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = ownerPassNum.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@ownerPassOVD@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = ownerPassOVD.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@ownerPassDate@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = ownerPassDate.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@ownerFullAddress@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = ownerAddress.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);




        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@ownerDoc@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = ownerDocs.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@registrationDoc@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = registrationDoc.Text;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@tehPass@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[41].Cells[1].Value.ToString(); ;

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.2@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[2].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.3@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[3].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.4@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[4].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.5@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[5].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.6@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[6].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.7@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[7].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.8@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[8].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.9@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[9].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.10@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[10].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.11@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[11].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.12@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[12].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.13@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[13].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.14@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[14].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.1.15@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[15].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);



        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.1@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[17].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.2@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[18].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.3@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[19].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.4@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[20].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.5@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[21].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.6@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[22].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.7@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[23].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.8@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[24].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.9@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[25].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.10@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[26].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.11@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[27].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.12@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[28].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.13@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[29].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.14@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[30].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.15@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[31].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.16@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[32].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);


        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.17@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[33].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.18@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[34].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.19@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[35].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.20@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[36].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.21@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[37].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.22@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[38].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.23@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[39].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.24@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[40].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.25@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[41].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.26@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[42].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.27@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[43].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.28@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[44].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.29@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[45].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.30@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[46].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.31@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[47].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.32@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[48].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.33@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[49].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.34@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[50].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.35@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[51].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.36@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[52].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.37@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[53].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.38@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[54].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.39@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[55].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.40@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[56].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);

        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "@@2.1.2.41@@";
        //                wdApp.Selection.Find.Replacement.ClearFormatting();
        //                wdApp.Selection.Find.Replacement.Text = dataGridView1.Rows[57].Cells[1].Value.ToString();

        //                wdApp.Selection.Find.Execute(
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                             ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);





        //                wdApp.Selection.Find.ClearFormatting();
        //                wdApp.Selection.Find.Text = "м2";
        //                while (wdApp.Selection.Find.Execute(
        //                                  ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                                  ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                                  ref Missing, ref Missing, ref Missing, ref Missing, ref Missing))
        //                {
        //                    wdApp.Selection.Characters[2].Font.Superscript = 1;
        //                }
        //                string te = wdApp.Selection.Text;
        //                //saving
        //                try
        //                {
        //                    int x = wdDoc.Shapes.Count;
        //                    x = wdDoc.Shapes.Count;
        //                    for (int k = 1; k < x; k++)
        //                    {
        //                        Microsoft.Office.Interop.Word.Shape shape = wdDoc.Shapes[k];

        //                        //string l = shape.AlternativeText;
        //                        if (shape.AlternativeText.Contains("cont"))
        //                        {
        //                            wdDoc.Shapes[k].TextEffect.Text = "№ " + contractNum.Text + " от " + calculationDate.Text + "г.";
        //                        }
        //                    }

        //                    //for (int k = 1; k < x; k++)
        //                    //{
        //                    //    Microsoft.Office.Interop.Word.Shape shape = wdDoc.Shapes[k];

        //                    //    if (shape.AlternativeText.Contains("first"))
        //                    //    {
        //                    //        System.Drawing.Image firstPageImg = System.Drawing.Image.FromFile(imagesGrid.Rows[0].Cells[2].Value.ToString());

        //                    //       //Clipboard.SetImage(firstPageImg);
        //                    //        shape.Select();
        //                    //        wdDoc.Shapes[k].CanvasItems.AddPicture(imagesGrid.Rows[0].Cells[2].Value.ToString());
        //                    //        wdDoc.Shapes[k].Apply();
        //                    //       // wdApp.Selection.PasteSpecial();
        //                    //       // wdApp.ActiveDocument.Shapes.AddPicture(imagesGrid.Rows[0].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing, Type.Missing, 500, 370, Type.Missing);
        //                    //        //wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[0].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                    //        //wdApp.Selection.InlineShapes
        //                    //        //Clipboard.Clear();
        //                    //    }

        //                    //}
        //                }
        //                catch (Exception exp)
        //                {

        //                }

        //                //foreach (Microsoft.Office.Interop.Word.Table table in wdApp.ActiveDocument.Tables)
        //                //{

        //                //    try
        //                //    {


        //                //        //  if (table.Columns[0].Cells[0].Range.Text.Contains("@@1@@"))
        //                //        //{
        //                //        foreach (Microsoft.Office.Interop.Word.Column col in table.Columns)
        //                //        {

        //                //            foreach (Microsoft.Office.Interop.Word.Cell cell in col.Cells)
        //                //            {
        //                //                int rowCount = imagesGrid.RowCount;
        //                //                string l = cell.Range.Text;
        //                //                if (l.Contains("@@1@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[1].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@1@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }

        //                //                if (l.Contains("@@2@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    //cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[2].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@2@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }
        //                //                if (l.Contains("@@3@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    //cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[3].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@3@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }

        //                //                if (l.Contains("@@4@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    //cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[4].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@4@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }

        //                //                if (l.Contains("@@5@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    //cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[5].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@5@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }

        //                //                if (l.Contains("@@6@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    //cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[6].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@6@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }

        //                //                if (l.Contains("@@7@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    //cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[7].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@7@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }

        //                //                if (l.Contains("@@8@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    //cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[8].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@8@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }

        //                //                if (l.Contains("@@9@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    //cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[9].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@9@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }

        //                //                if (l.Contains("@@10@@"))
        //                //                {
        //                //                    cell.Select();
        //                //                    //cell.Range.Text = "";
        //                //                    wdApp.Selection.InlineShapes.AddPicture(imagesGrid.Rows[10].Cells[2].Value.ToString(), Type.Missing, Type.Missing, Type.Missing);
        //                //                    wdApp.Selection.Find.ClearFormatting();
        //                //                    wdApp.Selection.Find.Text = "@@10@@";
        //                //                    wdApp.Selection.Find.Replacement.ClearFormatting();
        //                //                    wdApp.Selection.Find.Replacement.Text = "";

        //                //                    wdApp.Selection.Find.Execute(
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref Missing, ref Missing, ref Missing, ref Missing, ref Missing,
        //                //                                 ref replaceAll, ref Missing, ref Missing, ref Missing, ref Missing);
        //                //                }




        //                //            }
        //                //        }

        //                //    }
        //                //    //  }
        //                //    catch (Exception except)
        //                //    {
        //                //    }
        //                //}








        //                /*for (int k = 1; k < x; k++)
        //                {
        //                    Microsoft.Office.Interop.Word.Shape shape = wdDoc.Shapes[k];
        //                    float shift = 150;
        //                    string l = shape.AlternativeText;
        //                    if (l == "facade")
        //                    {
        //                        shape.IncrementTop(-shift);

        //                    }
        //                    if (l == "appartmentNum")
        //                    {
        //                        shape.IncrementTop(-shift);
        //                        //shape. = "Оцениваемая квартира №"+appartmentNum.Text;
        //                    }
        //                    if (l == "podezd")
        //                    {
        //                        shape.IncrementTop(-shift);
                        

        //                    }
        //                    if (l == "stairway")
        //                    {
        //                        shape.IncrementTop(-shift);

        //                    }
        //                }*/






        //                // 

        //                wdApp.ActiveDocument.SaveAs(saveFileDialog1.FileName);


        //                wdApp.Quit();
        //            }
        //        }
        //        catch (Exception except)
        //        {
        //            wdApp.Quit();
        //        }

            }
        }
    }
}
