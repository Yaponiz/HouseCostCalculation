using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System;
using System.Windows.Forms;
using HouseCostCalculation;

namespace WindowsFormsApplication1
{

    partial class mainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle11 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle16 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle17 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle18 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle19 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle20 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle23 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle24 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle25 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle21 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle22 = new System.Windows.Forms.DataGridViewCellStyle();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.dirtCalcPage = new System.Windows.Forms.TabPage();
            this.gridDoc2 = new System.Windows.Forms.TextBox();
            this.gridDoc = new System.Windows.Forms.TextBox();
            this.label47 = new System.Windows.Forms.Label();
            this.label48 = new System.Windows.Forms.Label();
            this.label41 = new System.Windows.Forms.Label();
            this.dirtKadastr = new System.Windows.Forms.TextBox();
            this.label40 = new System.Windows.Forms.Label();
            this.dirtm2 = new System.Windows.Forms.TextBox();
            this.saveGridToWordButton = new System.Windows.Forms.Button();
            this.dirtGridAnalogs = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn18 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn19 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn21 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button8 = new System.Windows.Forms.Button();
            this.dirtCalcGrid = new System.Windows.Forms.DataGridView();
            this.DirtName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.A1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.A2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.A3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.houseCalcPage = new System.Windows.Forms.TabPage();
            this.button3 = new System.Windows.Forms.Button();
            this.houseAnalogs = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn13 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn14 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn15 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn16 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn17 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.houseCalcGrid = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.housePage = new System.Windows.Forms.TabPage();
            this.button5 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.appartmentsCalcPage = new System.Windows.Forms.TabPage();
            this.saveAddsAppartaments = new System.Windows.Forms.Button();
            this.analogsGrid = new System.Windows.Forms.DataGridView();
            this.saveResultButton = new System.Windows.Forms.Button();
            this.saveAppartmentsCalc = new System.Windows.Forms.Button();
            this.calculationAppartaments = new System.Windows.Forms.DataGridView();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.appartmentsPage = new System.Windows.Forms.TabPage();
            this.objectDataGrid = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button2 = new System.Windows.Forms.Button();
            this.contractPage = new System.Windows.Forms.TabPage();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.SaveXMLButton = new System.Windows.Forms.Button();
            this.ownerOrg = new System.Windows.Forms.CheckBox();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.orgRegDate = new System.Windows.Forms.DateTimePicker();
            this.orgAdd = new System.Windows.Forms.TextBox();
            this.label38 = new System.Windows.Forms.Label();
            this.orgOGRN = new System.Windows.Forms.TextBox();
            this.orgKPP = new System.Windows.Forms.TextBox();
            this.orgINN = new System.Windows.Forms.TextBox();
            this.orgName = new System.Windows.Forms.TextBox();
            this.label42 = new System.Windows.Forms.Label();
            this.label43 = new System.Windows.Forms.Label();
            this.label44 = new System.Windows.Forms.Label();
            this.label45 = new System.Windows.Forms.Label();
            this.label46 = new System.Windows.Forms.Label();
            this.loadXMLButton = new System.Windows.Forms.Button();
            this.lm2text = new System.Windows.Forms.TextBox();
            this.ownerDocs = new System.Windows.Forms.TextBox();
            this.m2text = new System.Windows.Forms.TextBox();
            this.registrationDoc = new System.Windows.Forms.TextBox();
            this.label27 = new System.Windows.Forms.Label();
            this.label36 = new System.Windows.Forms.Label();
            this.label37 = new System.Windows.Forms.Label();
            this.label28 = new System.Windows.Forms.Label();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.gopage2 = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.newBuildingCheck = new System.Windows.Forms.CheckBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label39 = new System.Windows.Forms.Label();
            this.MO = new System.Windows.Forms.TextBox();
            this.Район = new System.Windows.Forms.Label();
            this.floors = new System.Windows.Forms.NumericUpDown();
            this.houseNum = new System.Windows.Forms.TextBox();
            this.label25 = new System.Windows.Forms.Label();
            this.roomsNum = new System.Windows.Forms.NumericUpDown();
            this.label24 = new System.Windows.Forms.Label();
            this.street = new System.Windows.Forms.ComboBox();
            this.lift = new System.Windows.Forms.ComboBox();
            this.houseType = new System.Windows.Forms.ComboBox();
            this.floor = new System.Windows.Forms.NumericUpDown();
            this.label22 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.appartmentNum = new System.Windows.Forms.TextBox();
            this.buildingNum = new System.Windows.Forms.TextBox();
            this.town = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label26 = new System.Windows.Forms.Label();
            this.bankName = new System.Windows.Forms.TextBox();
            this.calculationDate = new System.Windows.Forms.DateTimePicker();
            this.contractDate = new System.Windows.Forms.DateTimePicker();
            this.contractNum = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.showOwnersList = new System.Windows.Forms.Button();
            this.addOwner = new System.Windows.Forms.Button();
            this.ownerPassDate = new System.Windows.Forms.DateTimePicker();
            this.label32 = new System.Windows.Forms.Label();
            this.ownerPassport = new System.Windows.Forms.TextBox();
            this.label33 = new System.Windows.Forms.Label();
            this.label34 = new System.Windows.Forms.Label();
            this.label35 = new System.Windows.Forms.Label();
            this.ownerPassOVD = new System.Windows.Forms.TextBox();
            this.ownerPassNum = new System.Windows.Forms.TextBox();
            this.button10 = new System.Windows.Forms.Button();
            this.ownerPhone = new System.Windows.Forms.TextBox();
            this.ownerAddress = new System.Windows.Forms.TextBox();
            this.ownerInit = new System.Windows.Forms.TextBox();
            this.ownerName = new System.Windows.Forms.TextBox();
            this.ownerSurname = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.customerPassDate = new System.Windows.Forms.DateTimePicker();
            this.label31 = new System.Windows.Forms.Label();
            this.customerPassport = new System.Windows.Forms.TextBox();
            this.label30 = new System.Windows.Forms.Label();
            this.label29 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.customerPassOVD = new System.Windows.Forms.TextBox();
            this.customerPassNum = new System.Windows.Forms.TextBox();
            this.customerPadBut = new System.Windows.Forms.Button();
            this.ownerSameCustomer = new System.Windows.Forms.CheckBox();
            this.customerPhone = new System.Windows.Forms.TextBox();
            this.customerAddres = new System.Windows.Forms.TextBox();
            this.customerInit = new System.Windows.Forms.TextBox();
            this.customerName = new System.Windows.Forms.TextBox();
            this.customerSurname = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.form1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dirtCalcPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dirtGridAnalogs)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dirtCalcGrid)).BeginInit();
            this.houseCalcPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.houseAnalogs)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.houseCalcGrid)).BeginInit();
            this.housePage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.appartmentsCalcPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.analogsGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.calculationAppartaments)).BeginInit();
            this.appartmentsPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.objectDataGrid)).BeginInit();
            this.contractPage.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.floors)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.roomsNum)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.floor)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.form1BindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(196, 148);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // dirtCalcPage
            // 
            this.dirtCalcPage.Controls.Add(this.gridDoc2);
            this.dirtCalcPage.Controls.Add(this.gridDoc);
            this.dirtCalcPage.Controls.Add(this.label47);
            this.dirtCalcPage.Controls.Add(this.label48);
            this.dirtCalcPage.Controls.Add(this.label41);
            this.dirtCalcPage.Controls.Add(this.dirtKadastr);
            this.dirtCalcPage.Controls.Add(this.label40);
            this.dirtCalcPage.Controls.Add(this.dirtm2);
            this.dirtCalcPage.Controls.Add(this.saveGridToWordButton);
            this.dirtCalcPage.Controls.Add(this.dirtGridAnalogs);
            this.dirtCalcPage.Controls.Add(this.button8);
            this.dirtCalcPage.Controls.Add(this.dirtCalcGrid);
            this.dirtCalcPage.Location = new System.Drawing.Point(4, 22);
            this.dirtCalcPage.Name = "dirtCalcPage";
            this.dirtCalcPage.Size = new System.Drawing.Size(1134, 716);
            this.dirtCalcPage.TabIndex = 7;
            this.dirtCalcPage.Text = "Таблица оценки земли";
            this.dirtCalcPage.UseVisualStyleBackColor = true;
            // 
            // gridDoc2
            // 
            this.gridDoc2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.gridDoc2.Location = new System.Drawing.Point(577, 486);
            this.gridDoc2.Multiline = true;
            this.gridDoc2.Name = "gridDoc2";
            this.gridDoc2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.gridDoc2.Size = new System.Drawing.Size(361, 87);
            this.gridDoc2.TabIndex = 36;
            this.gridDoc2.Text = "Договор купли-продажи от 26/04/05г., зарегистрированный в Едином государственном " +
                "реестре прав за №15-15-01/026/2005-416; Акт приема-передачи от 26/04/05г.";
            // 
            // gridDoc
            // 
            this.gridDoc.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.gridDoc.Location = new System.Drawing.Point(577, 350);
            this.gridDoc.Multiline = true;
            this.gridDoc.Name = "gridDoc";
            this.gridDoc.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.gridDoc.Size = new System.Drawing.Size(361, 90);
            this.gridDoc.TabIndex = 34;
            this.gridDoc.Text = "Свидетельства о государственной регистрации права Управления Федеральной регистра" +
                "ционной службы по РСО-Алания серия 15 АЕ №689866 от 24/05/05г.";
            // 
            // label47
            // 
            this.label47.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label47.AutoSize = true;
            this.label47.Location = new System.Drawing.Point(577, 470);
            this.label47.Name = "label47";
            this.label47.Size = new System.Drawing.Size(193, 13);
            this.label47.TabIndex = 35;
            this.label47.Text = "Документы на право собственности";
            // 
            // label48
            // 
            this.label48.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label48.AutoSize = true;
            this.label48.Location = new System.Drawing.Point(577, 334);
            this.label48.Name = "label48";
            this.label48.Size = new System.Drawing.Size(180, 13);
            this.label48.TabIndex = 33;
            this.label48.Text = "Свидетельство о гос.регистрации";
            // 
            // label41
            // 
            this.label41.AutoSize = true;
            this.label41.Location = new System.Drawing.Point(627, 613);
            this.label41.Name = "label41";
            this.label41.Size = new System.Drawing.Size(110, 13);
            this.label41.TabIndex = 13;
            this.label41.Text = "Кадастровый номер";
            // 
            // dirtKadastr
            // 
            this.dirtKadastr.Location = new System.Drawing.Point(743, 613);
            this.dirtKadastr.Name = "dirtKadastr";
            this.dirtKadastr.Size = new System.Drawing.Size(284, 20);
            this.dirtKadastr.TabIndex = 12;
            // 
            // label40
            // 
            this.label40.AutoSize = true;
            this.label40.Location = new System.Drawing.Point(577, 579);
            this.label40.Name = "label40";
            this.label40.Size = new System.Drawing.Size(160, 13);
            this.label40.TabIndex = 11;
            this.label40.Text = "Площадь земельного участка";
            // 
            // dirtm2
            // 
            this.dirtm2.Location = new System.Drawing.Point(743, 579);
            this.dirtm2.Name = "dirtm2";
            this.dirtm2.Size = new System.Drawing.Size(100, 20);
            this.dirtm2.TabIndex = 10;
            this.dirtm2.TextChanged += new System.EventHandler(this.dirtm2_TextChanged_1);
            // 
            // saveGridToWordButton
            // 
            this.saveGridToWordButton.Location = new System.Drawing.Point(777, 654);
            this.saveGridToWordButton.Name = "saveGridToWordButton";
            this.saveGridToWordButton.Size = new System.Drawing.Size(173, 54);
            this.saveGridToWordButton.TabIndex = 8;
            this.saveGridToWordButton.Text = "Выгрузить отчет в Word";
            this.saveGridToWordButton.UseVisualStyleBackColor = true;
            this.saveGridToWordButton.Visible = false;
            this.saveGridToWordButton.Click += new System.EventHandler(this.saveGridToWordButton_Click);
            // 
            // dirtGridAnalogs
            // 
            this.dirtGridAnalogs.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dirtGridAnalogs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dirtGridAnalogs.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn18,
            this.dataGridViewTextBoxColumn19,
            this.dataGridViewTextBoxColumn20,
            this.dataGridViewTextBoxColumn21});
            this.dirtGridAnalogs.Location = new System.Drawing.Point(574, -2);
            this.dirtGridAnalogs.Name = "dirtGridAnalogs";
            this.dirtGridAnalogs.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dirtGridAnalogs.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dirtGridAnalogs.Size = new System.Drawing.Size(552, 313);
            this.dirtGridAnalogs.TabIndex = 7;
            this.dirtGridAnalogs.CellLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.dirtGridAnalogs_CellLeave);
            // 
            // dataGridViewTextBoxColumn18
            // 
            this.dataGridViewTextBoxColumn18.HeaderText = "Наименование";
            this.dataGridViewTextBoxColumn18.Name = "dataGridViewTextBoxColumn18";
            // 
            // dataGridViewTextBoxColumn19
            // 
            this.dataGridViewTextBoxColumn19.HeaderText = "Аналог №1";
            this.dataGridViewTextBoxColumn19.Name = "dataGridViewTextBoxColumn19";
            // 
            // dataGridViewTextBoxColumn20
            // 
            this.dataGridViewTextBoxColumn20.HeaderText = "Аналог №2";
            this.dataGridViewTextBoxColumn20.Name = "dataGridViewTextBoxColumn20";
            // 
            // dataGridViewTextBoxColumn21
            // 
            this.dataGridViewTextBoxColumn21.HeaderText = "Аналог №3";
            this.dataGridViewTextBoxColumn21.Name = "dataGridViewTextBoxColumn21";
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(574, 654);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(173, 54);
            this.button8.TabIndex = 5;
            this.button8.Text = "Сохранить рассчет";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // dirtCalcGrid
            // 
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dirtCalcGrid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dirtCalcGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dirtCalcGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DirtName,
            this.A1,
            this.A2,
            this.A3});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dirtCalcGrid.DefaultCellStyle = dataGridViewCellStyle3;
            this.dirtCalcGrid.Dock = System.Windows.Forms.DockStyle.Left;
            this.dirtCalcGrid.Location = new System.Drawing.Point(0, 0);
            this.dirtCalcGrid.Name = "dirtCalcGrid";
            this.dirtCalcGrid.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dirtCalcGrid.RowHeadersDefaultCellStyle = dataGridViewCellStyle4;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dirtCalcGrid.RowsDefaultCellStyle = dataGridViewCellStyle5;
            this.dirtCalcGrid.Size = new System.Drawing.Size(550, 716);
            this.dirtCalcGrid.TabIndex = 4;
            this.dirtCalcGrid.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dirtCalcGrid_CellEndEdit_1);
            this.dirtCalcGrid.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dirtCalcGrid_CellValueChanged);
            this.dirtCalcGrid.Enter += new System.EventHandler(this.dirtCalcGrid_Enter);
            // 
            // DirtName
            // 
            this.DirtName.HeaderText = "Название";
            this.DirtName.Name = "DirtName";
            // 
            // A1
            // 
            this.A1.HeaderText = "Аналог №1";
            this.A1.Name = "A1";
            // 
            // A2
            // 
            this.A2.HeaderText = "Аналог №2";
            this.A2.Name = "A2";
            // 
            // A3
            // 
            this.A3.HeaderText = "Аналог №3";
            this.A3.Name = "A3";
            // 
            // houseCalcPage
            // 
            this.houseCalcPage.Controls.Add(this.button3);
            this.houseCalcPage.Controls.Add(this.houseAnalogs);
            this.houseCalcPage.Controls.Add(this.button1);
            this.houseCalcPage.Controls.Add(this.button12);
            this.houseCalcPage.Controls.Add(this.houseCalcGrid);
            this.houseCalcPage.Location = new System.Drawing.Point(4, 22);
            this.houseCalcPage.Name = "houseCalcPage";
            this.houseCalcPage.Size = new System.Drawing.Size(1134, 716);
            this.houseCalcPage.TabIndex = 5;
            this.houseCalcPage.Text = "Таблица оценки домовладения";
            this.houseCalcPage.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(780, 639);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(173, 69);
            this.button3.TabIndex = 8;
            this.button3.Text = "Выгрузить приложение";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.saveAddsHouse_Click);
            // 
            // houseAnalogs
            // 
            this.houseAnalogs.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.houseAnalogs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.houseAnalogs.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn13,
            this.dataGridViewTextBoxColumn14,
            this.dataGridViewTextBoxColumn15,
            this.dataGridViewTextBoxColumn16,
            this.dataGridViewTextBoxColumn17});
            this.houseAnalogs.Location = new System.Drawing.Point(601, 3);
            this.houseAnalogs.Name = "houseAnalogs";
            this.houseAnalogs.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.houseAnalogs.RowsDefaultCellStyle = dataGridViewCellStyle6;
            this.houseAnalogs.Size = new System.Drawing.Size(533, 625);
            this.houseAnalogs.TabIndex = 7;
            this.houseAnalogs.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.houseAnalogs_CellEndEdit);
            // 
            // dataGridViewTextBoxColumn13
            // 
            this.dataGridViewTextBoxColumn13.HeaderText = "Наименование показателя";
            this.dataGridViewTextBoxColumn13.Name = "dataGridViewTextBoxColumn13";
            // 
            // dataGridViewTextBoxColumn14
            // 
            this.dataGridViewTextBoxColumn14.HeaderText = "Объект оценки";
            this.dataGridViewTextBoxColumn14.Name = "dataGridViewTextBoxColumn14";
            // 
            // dataGridViewTextBoxColumn15
            // 
            this.dataGridViewTextBoxColumn15.HeaderText = "Аналог № 1";
            this.dataGridViewTextBoxColumn15.Name = "dataGridViewTextBoxColumn15";
            // 
            // dataGridViewTextBoxColumn16
            // 
            this.dataGridViewTextBoxColumn16.HeaderText = "Аналог № 2";
            this.dataGridViewTextBoxColumn16.Name = "dataGridViewTextBoxColumn16";
            // 
            // dataGridViewTextBoxColumn17
            // 
            this.dataGridViewTextBoxColumn17.HeaderText = "Аналог № 3";
            this.dataGridViewTextBoxColumn17.Name = "dataGridViewTextBoxColumn17";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(601, 639);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(173, 69);
            this.button1.TabIndex = 6;
            this.button1.Text = "Сохранить рассчет";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(953, 639);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(173, 69);
            this.button12.TabIndex = 5;
            this.button12.Text = "Выгрузить в Word";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.SaveHouse);
            // 
            // houseCalcGrid
            // 
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.houseCalcGrid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
            this.houseCalcGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.houseCalcGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn8,
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10,
            this.dataGridViewTextBoxColumn11,
            this.dataGridViewTextBoxColumn12});
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.houseCalcGrid.DefaultCellStyle = dataGridViewCellStyle8;
            this.houseCalcGrid.Dock = System.Windows.Forms.DockStyle.Left;
            this.houseCalcGrid.Location = new System.Drawing.Point(0, 0);
            this.houseCalcGrid.Name = "houseCalcGrid";
            this.houseCalcGrid.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.houseCalcGrid.RowHeadersDefaultCellStyle = dataGridViewCellStyle9;
            dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.houseCalcGrid.RowsDefaultCellStyle = dataGridViewCellStyle10;
            this.houseCalcGrid.Size = new System.Drawing.Size(595, 716);
            this.houseCalcGrid.TabIndex = 4;
            this.houseCalcGrid.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.houseCalcGrid_CellEndEdit);
            this.houseCalcGrid.Enter += new System.EventHandler(this.houseCalcGrid_Enter);
            // 
            // dataGridViewTextBoxColumn8
            // 
            this.dataGridViewTextBoxColumn8.HeaderText = "Наименование показателя";
            this.dataGridViewTextBoxColumn8.Name = "dataGridViewTextBoxColumn8";
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.HeaderText = "Ед. изм.";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.HeaderText = "Аналог №1";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            // 
            // dataGridViewTextBoxColumn11
            // 
            this.dataGridViewTextBoxColumn11.HeaderText = "Аналог №2";
            this.dataGridViewTextBoxColumn11.Name = "dataGridViewTextBoxColumn11";
            // 
            // dataGridViewTextBoxColumn12
            // 
            this.dataGridViewTextBoxColumn12.HeaderText = "Аналог №3";
            this.dataGridViewTextBoxColumn12.Name = "dataGridViewTextBoxColumn12";
            // 
            // housePage
            // 
            this.housePage.Controls.Add(this.button5);
            this.housePage.Controls.Add(this.dataGridView1);
            this.housePage.Location = new System.Drawing.Point(4, 22);
            this.housePage.Name = "housePage";
            this.housePage.Size = new System.Drawing.Size(1134, 716);
            this.housePage.TabIndex = 4;
            this.housePage.Text = "Сведения о доме";
            this.housePage.UseVisualStyleBackColor = true;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(692, 488);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(214, 26);
            this.button5.TabIndex = 4;
            this.button5.Text = "Перейти к оценке";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            dataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle11;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.ColumnHeadersVisible = false;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6});
            dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle12;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Left;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle13;
            dataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowsDefaultCellStyle = dataGridViewCellStyle14;
            this.dataGridView1.Size = new System.Drawing.Size(445, 716);
            this.dataGridView1.TabIndex = 2;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.HeaderText = "Column1";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.HeaderText = "Column2";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            // 
            // appartmentsCalcPage
            // 
            this.appartmentsCalcPage.Controls.Add(this.saveAddsAppartaments);
            this.appartmentsCalcPage.Controls.Add(this.analogsGrid);
            this.appartmentsCalcPage.Controls.Add(this.saveResultButton);
            this.appartmentsCalcPage.Controls.Add(this.saveAppartmentsCalc);
            this.appartmentsCalcPage.Controls.Add(this.calculationAppartaments);
            this.appartmentsCalcPage.Location = new System.Drawing.Point(4, 22);
            this.appartmentsCalcPage.Name = "appartmentsCalcPage";
            this.appartmentsCalcPage.Padding = new System.Windows.Forms.Padding(3);
            this.appartmentsCalcPage.Size = new System.Drawing.Size(1134, 716);
            this.appartmentsCalcPage.TabIndex = 2;
            this.appartmentsCalcPage.Text = "Таблица оценки квартиры";
            this.appartmentsCalcPage.UseVisualStyleBackColor = true;
            // 
            // saveAddsAppartaments
            // 
            this.saveAddsAppartaments.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.saveAddsAppartaments.Location = new System.Drawing.Point(843, 634);
            this.saveAddsAppartaments.Name = "saveAddsAppartaments";
            this.saveAddsAppartaments.Size = new System.Drawing.Size(138, 76);
            this.saveAddsAppartaments.TabIndex = 7;
            this.saveAddsAppartaments.Text = "Выгрузить приложение";
            this.saveAddsAppartaments.UseVisualStyleBackColor = true;
            this.saveAddsAppartaments.Click += new System.EventHandler(this.saveAddsAppartaments_Click);
            // 
            // analogsGrid
            // 
            this.analogsGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.analogsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.analogsGrid.Location = new System.Drawing.Point(576, 3);
            this.analogsGrid.Name = "analogsGrid";
            this.analogsGrid.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.analogsGrid.RowsDefaultCellStyle = dataGridViewCellStyle15;
            this.analogsGrid.Size = new System.Drawing.Size(552, 625);
            this.analogsGrid.TabIndex = 6;
            this.analogsGrid.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.analogsGrid_CellValueChanged);
            // 
            // saveResultButton
            // 
            this.saveResultButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.saveResultButton.Location = new System.Drawing.Point(987, 634);
            this.saveResultButton.Name = "saveResultButton";
            this.saveResultButton.Size = new System.Drawing.Size(139, 76);
            this.saveResultButton.TabIndex = 5;
            this.saveResultButton.Text = "Выгрузить в Word";
            this.saveResultButton.UseVisualStyleBackColor = true;
            this.saveResultButton.Click += new System.EventHandler(this.saveResultButton_Click);
            // 
            // saveAppartmentsCalc
            // 
            this.saveAppartmentsCalc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.saveAppartmentsCalc.Location = new System.Drawing.Point(700, 634);
            this.saveAppartmentsCalc.Name = "saveAppartmentsCalc";
            this.saveAppartmentsCalc.Size = new System.Drawing.Size(137, 76);
            this.saveAppartmentsCalc.TabIndex = 3;
            this.saveAppartmentsCalc.Text = "Сохранить рассчет";
            this.saveAppartmentsCalc.UseVisualStyleBackColor = true;
            this.saveAppartmentsCalc.Click += new System.EventHandler(this.saveAppartmentsCalc_Click);
            // 
            // calculationAppartaments
            // 
            dataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.calculationAppartaments.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle16;
            this.calculationAppartaments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.calculationAppartaments.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6,
            this.Column7});
            dataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle17.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle17.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle17.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle17.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle17.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.calculationAppartaments.DefaultCellStyle = dataGridViewCellStyle17;
            this.calculationAppartaments.Dock = System.Windows.Forms.DockStyle.Left;
            this.calculationAppartaments.Location = new System.Drawing.Point(3, 3);
            this.calculationAppartaments.Name = "calculationAppartaments";
            this.calculationAppartaments.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle18.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle18.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle18.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle18.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle18.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.calculationAppartaments.RowHeadersDefaultCellStyle = dataGridViewCellStyle18;
            dataGridViewCellStyle19.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.calculationAppartaments.RowsDefaultCellStyle = dataGridViewCellStyle19;
            this.calculationAppartaments.Size = new System.Drawing.Size(567, 710);
            this.calculationAppartaments.TabIndex = 2;
            this.calculationAppartaments.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.calculationAppartaments_CellEndEdit);
            this.calculationAppartaments.CellLeave += new System.Windows.Forms.DataGridViewCellEventHandler(this.calculationAppartaments_CellLeave);
            this.calculationAppartaments.CellStateChanged += new System.Windows.Forms.DataGridViewCellStateChangedEventHandler(this.calculationAppartaments_CellStateChanged);
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Наименование показателя";
            this.Column3.Name = "Column3";
            // 
            // Column4
            // 
            this.Column4.HeaderText = "Ед. изм.";
            this.Column4.Name = "Column4";
            // 
            // Column5
            // 
            this.Column5.HeaderText = "Аналог №1";
            this.Column5.Name = "Column5";
            // 
            // Column6
            // 
            this.Column6.HeaderText = "Аналог №2";
            this.Column6.Name = "Column6";
            // 
            // Column7
            // 
            this.Column7.HeaderText = "Аналог №3";
            this.Column7.Name = "Column7";
            // 
            // appartmentsPage
            // 
            this.appartmentsPage.Controls.Add(this.objectDataGrid);
            this.appartmentsPage.Controls.Add(this.button2);
            this.appartmentsPage.Location = new System.Drawing.Point(4, 22);
            this.appartmentsPage.Name = "appartmentsPage";
            this.appartmentsPage.Padding = new System.Windows.Forms.Padding(3);
            this.appartmentsPage.Size = new System.Drawing.Size(1134, 716);
            this.appartmentsPage.TabIndex = 1;
            this.appartmentsPage.Text = "Cведения о квартире";
            this.appartmentsPage.UseVisualStyleBackColor = true;
            // 
            // objectDataGrid
            // 
            dataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle20.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle20.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle20.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle20.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle20.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.objectDataGrid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle20;
            this.objectDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.objectDataGrid.ColumnHeadersVisible = false;
            this.objectDataGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2});
            dataGridViewCellStyle23.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle23.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle23.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle23.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle23.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle23.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle23.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.objectDataGrid.DefaultCellStyle = dataGridViewCellStyle23;
            this.objectDataGrid.Dock = System.Windows.Forms.DockStyle.Left;
            this.objectDataGrid.Location = new System.Drawing.Point(3, 3);
            this.objectDataGrid.Name = "objectDataGrid";
            this.objectDataGrid.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle24.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle24.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle24.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle24.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle24.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle24.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle24.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.objectDataGrid.RowHeadersDefaultCellStyle = dataGridViewCellStyle24;
            dataGridViewCellStyle25.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.objectDataGrid.RowsDefaultCellStyle = dataGridViewCellStyle25;
            this.objectDataGrid.Size = new System.Drawing.Size(698, 710);
            this.objectDataGrid.TabIndex = 1;
            this.objectDataGrid.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.objectDataGrid_CellEndEdit);
            // 
            // Column1
            // 
            dataGridViewCellStyle21.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Column1.DefaultCellStyle = dataGridViewCellStyle21;
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            this.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Column2
            // 
            dataGridViewCellStyle22.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Column2.DefaultCellStyle = dataGridViewCellStyle22;
            this.Column2.HeaderText = "Column2";
            this.Column2.Name = "Column2";
            this.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Column2.Width = 500;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(794, 645);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(214, 26);
            this.button2.TabIndex = 0;
            this.button2.Text = "Перейти к оценке";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // contractPage
            // 
            this.contractPage.Controls.Add(this.checkBox1);
            this.contractPage.Controls.Add(this.SaveXMLButton);
            this.contractPage.Controls.Add(this.ownerOrg);
            this.contractPage.Controls.Add(this.groupBox9);
            this.contractPage.Controls.Add(this.loadXMLButton);
            this.contractPage.Controls.Add(this.lm2text);
            this.contractPage.Controls.Add(this.ownerDocs);
            this.contractPage.Controls.Add(this.m2text);
            this.contractPage.Controls.Add(this.registrationDoc);
            this.contractPage.Controls.Add(this.label27);
            this.contractPage.Controls.Add(this.label36);
            this.contractPage.Controls.Add(this.label37);
            this.contractPage.Controls.Add(this.label28);
            this.contractPage.Controls.Add(this.groupBox6);
            this.contractPage.Controls.Add(this.gopage2);
            this.contractPage.Controls.Add(this.groupBox4);
            this.contractPage.Controls.Add(this.groupBox3);
            this.contractPage.Controls.Add(this.groupBox2);
            this.contractPage.Controls.Add(this.groupBox1);
            this.contractPage.Location = new System.Drawing.Point(4, 22);
            this.contractPage.Name = "contractPage";
            this.contractPage.Padding = new System.Windows.Forms.Padding(3);
            this.contractPage.Size = new System.Drawing.Size(1134, 716);
            this.contractPage.TabIndex = 0;
            this.contractPage.Text = "Сведения для договора";
            this.contractPage.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(467, 493);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(74, 17);
            this.checkBox1.TabIndex = 39;
            this.checkBox1.Text = "Заказчик";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // SaveXMLButton
            // 
            this.SaveXMLButton.Location = new System.Drawing.Point(767, 668);
            this.SaveXMLButton.Name = "SaveXMLButton";
            this.SaveXMLButton.Size = new System.Drawing.Size(93, 23);
            this.SaveXMLButton.TabIndex = 38;
            this.SaveXMLButton.Text = "Сохранить";
            this.SaveXMLButton.UseVisualStyleBackColor = true;
            this.SaveXMLButton.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // ownerOrg
            // 
            this.ownerOrg.AutoSize = true;
            this.ownerOrg.Location = new System.Drawing.Point(352, 492);
            this.ownerOrg.Name = "ownerOrg";
            this.ownerOrg.Size = new System.Drawing.Size(92, 17);
            this.ownerOrg.TabIndex = 37;
            this.ownerOrg.Text = "Собственник";
            this.ownerOrg.UseVisualStyleBackColor = true;
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.orgRegDate);
            this.groupBox9.Controls.Add(this.orgAdd);
            this.groupBox9.Controls.Add(this.label38);
            this.groupBox9.Controls.Add(this.orgOGRN);
            this.groupBox9.Controls.Add(this.orgKPP);
            this.groupBox9.Controls.Add(this.orgINN);
            this.groupBox9.Controls.Add(this.orgName);
            this.groupBox9.Controls.Add(this.label42);
            this.groupBox9.Controls.Add(this.label43);
            this.groupBox9.Controls.Add(this.label44);
            this.groupBox9.Controls.Add(this.label45);
            this.groupBox9.Controls.Add(this.label46);
            this.groupBox9.Location = new System.Drawing.Point(346, 509);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.groupBox9.Size = new System.Drawing.Size(327, 172);
            this.groupBox9.TabIndex = 35;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Организация";
            // 
            // orgRegDate
            // 
            this.orgRegDate.CustomFormat = "dd/MM/yy";
            this.orgRegDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.orgRegDate.Location = new System.Drawing.Point(108, 120);
            this.orgRegDate.Name = "orgRegDate";
            this.orgRegDate.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.orgRegDate.Size = new System.Drawing.Size(204, 20);
            this.orgRegDate.TabIndex = 14;
            // 
            // orgAdd
            // 
            this.orgAdd.Location = new System.Drawing.Point(109, 146);
            this.orgAdd.Name = "orgAdd";
            this.orgAdd.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.orgAdd.Size = new System.Drawing.Size(204, 20);
            this.orgAdd.TabIndex = 11;
            // 
            // label38
            // 
            this.label38.AutoSize = true;
            this.label38.Location = new System.Drawing.Point(6, 146);
            this.label38.Name = "label38";
            this.label38.Size = new System.Drawing.Size(38, 13);
            this.label38.TabIndex = 10;
            this.label38.Text = "Адрес";
            // 
            // orgOGRN
            // 
            this.orgOGRN.Location = new System.Drawing.Point(108, 93);
            this.orgOGRN.Name = "orgOGRN";
            this.orgOGRN.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.orgOGRN.Size = new System.Drawing.Size(204, 20);
            this.orgOGRN.TabIndex = 8;
            // 
            // orgKPP
            // 
            this.orgKPP.Location = new System.Drawing.Point(108, 68);
            this.orgKPP.Name = "orgKPP";
            this.orgKPP.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.orgKPP.Size = new System.Drawing.Size(204, 20);
            this.orgKPP.TabIndex = 7;
            // 
            // orgINN
            // 
            this.orgINN.Location = new System.Drawing.Point(108, 42);
            this.orgINN.Name = "orgINN";
            this.orgINN.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.orgINN.Size = new System.Drawing.Size(204, 20);
            this.orgINN.TabIndex = 6;
            // 
            // orgName
            // 
            this.orgName.Location = new System.Drawing.Point(108, 13);
            this.orgName.Name = "orgName";
            this.orgName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.orgName.Size = new System.Drawing.Size(204, 20);
            this.orgName.TabIndex = 5;
            // 
            // label42
            // 
            this.label42.AutoSize = true;
            this.label42.Location = new System.Drawing.Point(5, 119);
            this.label42.Name = "label42";
            this.label42.Size = new System.Drawing.Size(100, 13);
            this.label42.TabIndex = 4;
            this.label42.Text = "Дата регистрации";
            // 
            // label43
            // 
            this.label43.AutoSize = true;
            this.label43.Location = new System.Drawing.Point(5, 93);
            this.label43.Name = "label43";
            this.label43.Size = new System.Drawing.Size(36, 13);
            this.label43.TabIndex = 3;
            this.label43.Text = "ОГРН";
            // 
            // label44
            // 
            this.label44.AutoSize = true;
            this.label44.Location = new System.Drawing.Point(5, 67);
            this.label44.Name = "label44";
            this.label44.Size = new System.Drawing.Size(30, 13);
            this.label44.TabIndex = 2;
            this.label44.Text = "КПП";
            // 
            // label45
            // 
            this.label45.AutoSize = true;
            this.label45.Location = new System.Drawing.Point(5, 42);
            this.label45.Name = "label45";
            this.label45.Size = new System.Drawing.Size(31, 13);
            this.label45.TabIndex = 1;
            this.label45.Text = "ИНН";
            // 
            // label46
            // 
            this.label46.AutoSize = true;
            this.label46.Location = new System.Drawing.Point(5, 16);
            this.label46.Name = "label46";
            this.label46.Size = new System.Drawing.Size(83, 13);
            this.label46.TabIndex = 0;
            this.label46.Text = "Наименование";
            // 
            // loadXMLButton
            // 
            this.loadXMLButton.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.loadXMLButton.Location = new System.Drawing.Point(767, 637);
            this.loadXMLButton.Name = "loadXMLButton";
            this.loadXMLButton.Size = new System.Drawing.Size(93, 25);
            this.loadXMLButton.TabIndex = 34;
            this.loadXMLButton.Text = "Загрузить";
            this.loadXMLButton.UseVisualStyleBackColor = true;
            this.loadXMLButton.Click += new System.EventHandler(this.loadDataButton);
            // 
            // lm2text
            // 
            this.lm2text.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.lm2text.Location = new System.Drawing.Point(913, 343);
            this.lm2text.Name = "lm2text";
            this.lm2text.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lm2text.Size = new System.Drawing.Size(183, 20);
            this.lm2text.TabIndex = 33;
            this.lm2text.TextChanged += new System.EventHandler(this.lm2text_TextChanged);
            // 
            // ownerDocs
            // 
            this.ownerDocs.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.ownerDocs.Location = new System.Drawing.Point(767, 530);
            this.ownerDocs.Multiline = true;
            this.ownerDocs.Name = "ownerDocs";
            this.ownerDocs.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerDocs.Size = new System.Drawing.Size(361, 87);
            this.ownerDocs.TabIndex = 32;
            this.ownerDocs.Text = "Договор купли-продажи от 26/04/05г., зарегистрированный в Едином государственном " +
                "реестре прав за №15-15-01/026/2005-416; Акт приема-передачи от 26/04/05г.";
            this.ownerDocs.TextChanged += new System.EventHandler(this.ownerDocs_TextChanged);
            // 
            // m2text
            // 
            this.m2text.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.m2text.Location = new System.Drawing.Point(913, 317);
            this.m2text.Name = "m2text";
            this.m2text.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.m2text.Size = new System.Drawing.Size(183, 20);
            this.m2text.TabIndex = 32;
            this.m2text.TextChanged += new System.EventHandler(this.m2text_TextChanged);
            // 
            // registrationDoc
            // 
            this.registrationDoc.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.registrationDoc.Location = new System.Drawing.Point(767, 394);
            this.registrationDoc.Multiline = true;
            this.registrationDoc.Name = "registrationDoc";
            this.registrationDoc.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.registrationDoc.Size = new System.Drawing.Size(361, 90);
            this.registrationDoc.TabIndex = 29;
            this.registrationDoc.Text = "Свидетельства о государственной регистрации права Управления Федеральной регистра" +
                "ционной службы по РСО-Алания серия 15 АЕ №689866 от 24/05/05г.";
            this.registrationDoc.TextChanged += new System.EventHandler(this.registrationDoc_TextChanged);
            // 
            // label27
            // 
            this.label27.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(810, 343);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(90, 13);
            this.label27.TabIndex = 31;
            this.label27.Text = "Жилая площадь";
            // 
            // label36
            // 
            this.label36.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label36.AutoSize = true;
            this.label36.Location = new System.Drawing.Point(767, 514);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(193, 13);
            this.label36.TabIndex = 31;
            this.label36.Text = "Документы на право собственности";
            // 
            // label37
            // 
            this.label37.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label37.AutoSize = true;
            this.label37.Location = new System.Drawing.Point(810, 317);
            this.label37.Name = "label37";
            this.label37.Size = new System.Drawing.Size(90, 13);
            this.label37.TabIndex = 30;
            this.label37.Text = "Общая площадь";
            // 
            // label28
            // 
            this.label28.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label28.AutoSize = true;
            this.label28.Location = new System.Drawing.Point(767, 378);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(180, 13);
            this.label28.TabIndex = 16;
            this.label28.Text = "Свидетельство о гос.регистрации";
            // 
            // groupBox6
            // 
            this.groupBox6.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.groupBox6.Controls.Add(this.pictureBox1);
            this.groupBox6.Location = new System.Drawing.Point(820, 16);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(308, 278);
            this.groupBox6.TabIndex = 13;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Карта";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::HouseCostCalculation.Properties.Resources.map_Vladikavkaz;
            this.pictureBox1.ImageLocation = "";
            this.pictureBox1.Location = new System.Drawing.Point(24, 19);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(265, 241);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // gopage2
            // 
            this.gopage2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.gopage2.Location = new System.Drawing.Point(913, 637);
            this.gopage2.Name = "gopage2";
            this.gopage2.Size = new System.Drawing.Size(215, 66);
            this.gopage2.TabIndex = 30;
            this.gopage2.Text = "Перейти к сведениям о квартире";
            this.gopage2.UseVisualStyleBackColor = true;
            this.gopage2.Click += new System.EventHandler(this.button3_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.newBuildingCheck);
            this.groupBox4.Controls.Add(this.textBox1);
            this.groupBox4.Controls.Add(this.label39);
            this.groupBox4.Controls.Add(this.MO);
            this.groupBox4.Controls.Add(this.Район);
            this.groupBox4.Controls.Add(this.floors);
            this.groupBox4.Controls.Add(this.houseNum);
            this.groupBox4.Controls.Add(this.label25);
            this.groupBox4.Controls.Add(this.roomsNum);
            this.groupBox4.Controls.Add(this.label24);
            this.groupBox4.Controls.Add(this.street);
            this.groupBox4.Controls.Add(this.lift);
            this.groupBox4.Controls.Add(this.houseType);
            this.groupBox4.Controls.Add(this.floor);
            this.groupBox4.Controls.Add(this.label22);
            this.groupBox4.Controls.Add(this.label21);
            this.groupBox4.Controls.Add(this.label20);
            this.groupBox4.Controls.Add(this.label19);
            this.groupBox4.Controls.Add(this.textBox11);
            this.groupBox4.Controls.Add(this.appartmentNum);
            this.groupBox4.Controls.Add(this.buildingNum);
            this.groupBox4.Controls.Add(this.town);
            this.groupBox4.Controls.Add(this.label11);
            this.groupBox4.Controls.Add(this.label12);
            this.groupBox4.Controls.Add(this.label16);
            this.groupBox4.Controls.Add(this.label17);
            this.groupBox4.Controls.Add(this.label18);
            this.groupBox4.Location = new System.Drawing.Point(336, 3);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(478, 348);
            this.groupBox4.TabIndex = 11;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Оцениваемая собственность";
            // 
            // newBuildingCheck
            // 
            this.newBuildingCheck.AutoSize = true;
            this.newBuildingCheck.Location = new System.Drawing.Point(309, 69);
            this.newBuildingCheck.Name = "newBuildingCheck";
            this.newBuildingCheck.Size = new System.Drawing.Size(93, 17);
            this.newBuildingCheck.TabIndex = 38;
            this.newBuildingCheck.Text = "Новостройка";
            this.newBuildingCheck.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(309, 39);
            this.textBox1.Name = "textBox1";
            this.textBox1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.textBox1.Size = new System.Drawing.Size(162, 20);
            this.textBox1.TabIndex = 31;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label39
            // 
            this.label39.AutoSize = true;
            this.label39.Location = new System.Drawing.Point(309, 19);
            this.label39.Name = "label39";
            this.label39.Size = new System.Drawing.Size(100, 13);
            this.label39.TabIndex = 30;
            this.label39.Text = "Уточнение района";
            // 
            // MO
            // 
            this.MO.Location = new System.Drawing.Point(108, 315);
            this.MO.Name = "MO";
            this.MO.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.MO.Size = new System.Drawing.Size(183, 20);
            this.MO.TabIndex = 29;
            this.MO.TextChanged += new System.EventHandler(this.MO_TextChanged);
            // 
            // Район
            // 
            this.Район.AutoSize = true;
            this.Район.Location = new System.Drawing.Point(5, 315);
            this.Район.Name = "Район";
            this.Район.Size = new System.Drawing.Size(38, 13);
            this.Район.TabIndex = 28;
            this.Район.Text = "Район";
            // 
            // floors
            // 
            this.floors.Location = new System.Drawing.Point(108, 230);
            this.floors.Name = "floors";
            this.floors.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.floors.Size = new System.Drawing.Size(183, 20);
            this.floors.TabIndex = 20;
            this.floors.ValueChanged += new System.EventHandler(this.floors_ValueChanged);
            // 
            // houseNum
            // 
            this.houseNum.Location = new System.Drawing.Point(108, 69);
            this.houseNum.Name = "houseNum";
            this.houseNum.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.houseNum.Size = new System.Drawing.Size(183, 20);
            this.houseNum.TabIndex = 7;
            this.houseNum.TextChanged += new System.EventHandler(this.houseNum_TextChanged);
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(5, 69);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(30, 13);
            this.label25.TabIndex = 25;
            this.label25.Text = "Дом";
            // 
            // roomsNum
            // 
            this.roomsNum.Location = new System.Drawing.Point(108, 287);
            this.roomsNum.Name = "roomsNum";
            this.roomsNum.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.roomsNum.Size = new System.Drawing.Size(183, 20);
            this.roomsNum.TabIndex = 22;
            this.roomsNum.ValueChanged += new System.EventHandler(this.roomsNum_ValueChanged);
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(6, 289);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(81, 13);
            this.label24.TabIndex = 23;
            this.label24.Text = "Кол-во комнат";
            // 
            // street
            // 
            this.street.FormattingEnabled = true;
            this.street.Items.AddRange(new object[] {
            "ул. 1 Мая",
            "ул. 40 лет Победы",
            "Площадь 50-летия Октября",
            "ул. 8 Марта",
            "ул. Августовских Событий",
            "Автобусный переулок",
            "Алагирская улица",
            "ул. Алибека Кантемирова",
            "Апшеронский переулок",
            "Ардонская улица",
            "Армянская улица",
            "Артиллерийская улица",
            "Архонский проезд",
            "ул. Астана Кесаева",
            "Базарный переулок",
            "Бакинская улица",
            "Балкинский проезд",
            "ул. Баракова (Шалдонская)",
            "ул. Барбашова",
            "Батумская улица",
            "Безымянный переулок",
            "ул. Белинского",
            "Беляевский переулок",
            "Бесланская улица",
            "ул. Бзарова",
            "ул. Бибо Ватаева",
            "Бородинская улица",
            "ул. Борукаева",
            "ул. Ботоева",
            "Братская улица",
            "ул. Братьев Габайраевых",
            "ул. Братьев Газдановых",
            "ул. Братьев Темировых",
            "ул. Братьев Щукиных",
            "Брестская улица",
            "ул. Бритаева",
            "Бульварная улица",
            "ул. Бутаева",
            "ул. Бутырина",
            "ул. Вадима Эльмесова",
            "ул. Васо Баева",
            "ул. Ватутина",
            "ул. Вахтангова",
            "Верхняя улица",
            "Веселая улица",
            "Весенняя улица",
            "Виноградный переулок",
            "Вишневый переулок",
            "Владикавказская улица",
            "ул. Владимира Тхапсаева",
            "Военный 29-й городок",
            "ул. Войкова",
            "Волгоградская улица",
            "Волжская улица",
            "ул. Воробьева",
            "Восточный переулок",
            "ул. Гагарина",
            "ул. Гадиева",
            "ул. Гайдара",
            "ул. Гаппо Баева (Свободы)",
            "ул. Гастелло",
            "Гвардейская улица",
            "ул. Гегечкори",
            "ул. Генерала Дзусова",
            "ул. Генерала Плиева",
            "ул. Герасимова",
            "ул. Герцена",
            "ул. Гибизова",
            "Гизельский переулок",
            "Гизельское шоссе",
            "ул. Гикало",
            "ул. Гоголя",
            "ул. Годовикова",
            "ул. Гончарова",
            "ул. Гостиева",
            "ул. Грибоедова",
            "Грозненская улица",
            "Грузинская улица",
            "ул. Гугкаева",
            "ул. Гудованцева",
            "Гэсовская улица",
            "ул. Д.Донского",
            "Даргавский переулок",
            "Дарьяльская улица",
            "Дачная улица",
            "ул. Декабристов",
            "Дербентская улица",
            "ул. Дзарахохова",
            "ул. Дзержинского",
            "ул. Дивизии НКВД (Редантс)",
            "Дигорская улица",
            "ул. Димитрова",
            "Длинно-Долинская улица",
            "ул. Добролюбова",
            "Проспект Доватора",
            "ул. Дружбы",
            "Ереванский переулок",
            "ул. Есенина (Кадгаронская)",
            "Железнодорожный переулок",
            "Заводская улица",
            "Заводской переулок",
            "Загородная улица",
            "ул. Зангиева",
            "Зеленая улица",
            "ул. Земнухова",
            "Зильгинский переулок",
            "ул. Зои Космодемьянской",
            "ул. Зортова",
            "ул. Зураба Магкаева",
            "Интернациональная улица",
            "Иронский переулок",
            "Кабардинская улица",
            "ул. Камалова",
            "ул. Камбердиева",
            "ул. Кантемирова",
            "Карджинский переулок",
            "ул. Карла Маркса",
            "Карцинское шоссе",
            "Керамический переулок",
            "ул. Керменистов",
            "Кирпичный переулок",
            "Кисловодская улица",
            "Ключевская улица",
            "Кобинский переулок",
            "ул. Коблова",
            "Ковровая улица",
            "Кожевенный переулок",
            "ул. Койбаева",
            "ул. Кольбуса",
            "Переулок Кольцова",
            "ул. Кольцова",
            "ул. Коммунаров",
            "Комсомольская улица",
            "ул. Корнеева",
            "ул. Коста Хетагурова",
            "ул. Котовского",
            "ул. Коцоева",
            "Крайняя улица",
            "Красная улица",
            "Красноармейская улица",
            "Краснодонская улица",
            "Краснорядская площадь (бывш. площадь Штыба)",
            "Кривой переулок",
            "Крупская улица",
            "ул. Крылова",
            "Крымская улица",
            "Крымский переулок",
            "ул. Кубалова",
            "ул. Куйбышева",
            "Курганная улица",
            "Курортная улица",
            "Курская улица",
            "ул. Кутузова",
            "ул. Кцоева",
            "Кырджалийская улица",
            "ул. Ларионова",
            "Ларская улица",
            "ул. Левандовского",
            "ул. Левитана",
            "Площадь Ленина",
            "ул. Ленина",
            "ул. Леонова",
            "Лермонтовская улица",
            "Лесная улица",
            "ул. Лобачевского",
            "Луговая улица",
            "ул. Льва Толстого",
            "ул. Любови Шевцовой",
            "ул. Макаренко",
            "ул. Максима Горького",
            "Малая улица",
            "Малгобекская улица",
            "ул. Малиева",
            "Малый переулок",
            "Мамисонский переулок",
            "ул. Мамсурова",
            "ул. Мамсурова Хаджи",
            "ул. Марины Расковой",
            "ул. Г. Масленникова",
            "ул. Матросова",
            "ул. Маяковского",
            "Межевая улица",
            "ул. Менделеева",
            "ул. Металлургов",
            "ул. Минина",
            "Проспект Мира",
            "ул. Митькина",
            "ул. Мичурина",
            "Моздокская улица",
            "Молодежная улица",
            "Молодежный переулок",
            "ул. Мордовцева",
            "ул. Морских пехотинцев",
            "Московская улица",
            "Московское шоссе",
            "Музейный переулок",
            "Нагорная улица",
            "Нальчикская улица",
            "ул. Народов Востока",
            "Нартовская улица",
            "ул. Нахимова",
            "ул. Неведомского",
            "Невский переулок",
            "ул. Неизвестного Солдата",
            "ул. Некрасова",
            "ул. Никитина",
            "ул. Николаева",
            "Новагинская улица",
            "Новая улица",
            "Ногирская улица",
            "Обрывистая улица",
            "ул. Огнева",
            "ул. Огурцова",
            "Озёрная улица",
            "ул. Олега Кошевого",
            "Оружейная улица",
            "Осетинская улица",
            "Переулок Осипенко",
            "ул. Остаева",
            "ул. Островского",
            "Охотничий переулок",
            "ул. Павленко",
            "ул. Павлика Морозова",
            "Павловский переулок",
            "Партизанский переулок",
            "ул. Пашковского",
            "Первомайская улица",
            "Петровский переулок",
            "ул. Пионеров",
            "ул. Пироговского",
            "Площадь Победы",
            "ул. Побежимова",
            "Пограничная улица",
            "ул. Пожарского",
            "Покровский переулок",
            "ул. Попова",
            "Почтовая улица",
            "Предмостный переулок",
            "Пригородная улица",
            "Придорожная улица",
            "Продуктовый переулок",
            "Промышленная 1-я улица",
            "Промышленная 3-я улица",
            "Промышленная 4-я улица",
            "Промышленная 5-я улица",
            "Промышленная 6-я улица",
            "Промышленная 7-я улица",
            "Пугачевский переулок",
            "Сквер Пушкина",
            "Пушкинская улица",
            "Пчеловодная улица",
            "Рабочий переулок",
            "ул. Рамонова",
            "ул. Революции",
            "Речная улица",
            "Рыночный переулок",
            "Садовая улица",
            "Садонская улица",
            "ул. Сады Шалдона",
            "Санаторный переулок",
            "Свердловская улица",
            "Светлая улица",
            "Площадь Свободы",
            "Севастопольская улица",
            "Северная улица",
            "ул. Седова",
            "Сельская улица",
            "ул. Серафимовича",
            "ул. Серобабова",
            "ул. Серова",
            "Сибирская улица",
            "Сквозная улица",
            "Слободской переулок",
            "Слюсаревский переулок",
            "Солнечная улица",
            "Соляный переулок",
            "ул. Спартака",
            "Спокойная улица",
            "Спортивная улица",
            "Средняя улица",
            "Ставропольская улица",
            "Переулок Станиславского",
            "ул. Стаханова",
            "ул. Степана Разина",
            "Суворовская улица",
            "ул. Таболова",
            "ул. Талалихина",
            "ул. Тамаева",
            "Тарская улица",
            "Тарское шоссе",
            "ул. Таутиева",
            "Театральный переулок",
            "ул. Тельмана",
            "Тимирязевский переулок",
            "ул. Титова",
            "ул. Тихий",
            "ул. Тогоева",
            "Торговый переулок",
            "ул. Торчинова",
            "Тракторный переулок",
            "Транспортный переулок",
            "ул. Триандофилова",
            "ул. Трубецкого",
            "Трудовая улица",
            "Тупиковый переулок",
            "Турбинная улица",
            "ул. Уруймаговой",
            "ул. Езетхан Уруймаговой",
            "Уфимская улица",
            "ул. Ушакова",
            "ул. Ушинского",
            "ул. Фрунзе",
            "Хазнидонская улица",
            "Холодный переулок",
            "ул. Цаголова",
            "ул. Цаликова",
            "Цейская улица",
            "Целинная улица",
            "ул. Церетели",
            "ул. Цоколаева",
            "ул. Чайковского",
            "ул. Чапаева",
            "ул. Чермена Баева",
            "Черменский проезд",
            "Черменское шоссе",
            "ул. Черноглаза",
            "Черноморская улица",
            "ул. Черняховского",
            "ул. Чехова",
            "ул. Чкалова",
            "ул. Шевченко",
            "ул. Шёгрена",
            "Школьный переулок",
            "ул. Шмулевича",
            "Шоссейная улица",
            "ул. Шота Руставели",
            "ул. Штыба",
            "ул. Щербакова",
            "Южная улица",
            "Ягодная улица",
            "ул. Яшина"});
            this.street.Location = new System.Drawing.Point(108, 42);
            this.street.Name = "street";
            this.street.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.street.Size = new System.Drawing.Size(183, 21);
            this.street.TabIndex = 6;
            this.street.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectedIndexChanged);
            this.street.TextChanged += new System.EventHandler(this.street_TextChanged);
            this.street.KeyUp += new System.Windows.Forms.KeyEventHandler(this.street_KeyUp);
            // 
            // lift
            // 
            this.lift.FormattingEnabled = true;
            this.lift.Items.AddRange(new object[] {
            "Есть",
            "Нет"});
            this.lift.Location = new System.Drawing.Point(108, 203);
            this.lift.Name = "lift";
            this.lift.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lift.Size = new System.Drawing.Size(183, 21);
            this.lift.TabIndex = 19;
            this.lift.SelectedIndexChanged += new System.EventHandler(this.lift_SelectedIndexChanged);
            // 
            // houseType
            // 
            this.houseType.FormattingEnabled = true;
            this.houseType.Items.AddRange(new object[] {
            "Кирпичный",
            "Панельный",
            "Монолитный"});
            this.houseType.Location = new System.Drawing.Point(108, 257);
            this.houseType.Name = "houseType";
            this.houseType.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.houseType.Size = new System.Drawing.Size(183, 21);
            this.houseType.TabIndex = 21;
            this.houseType.SelectedIndexChanged += new System.EventHandler(this.registrationDoc_TextChanged);
            // 
            // floor
            // 
            this.floor.Location = new System.Drawing.Point(108, 177);
            this.floor.Name = "floor";
            this.floor.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.floor.Size = new System.Drawing.Size(183, 20);
            this.floor.TabIndex = 18;
            this.floor.ValueChanged += new System.EventHandler(this.numericUpDown1_ValueChanged);
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(5, 257);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(55, 13);
            this.label22.TabIndex = 16;
            this.label22.Text = "Тип дома";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(5, 231);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(91, 13);
            this.label21.TabIndex = 14;
            this.label21.Text = "Этажность дома";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(5, 205);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(34, 13);
            this.label20.TabIndex = 12;
            this.label20.Text = "Лифт";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(5, 177);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(33, 13);
            this.label19.TabIndex = 10;
            this.label19.Text = "Этаж";
            // 
            // textBox11
            // 
            this.textBox11.Location = new System.Drawing.Point(108, 151);
            this.textBox11.Name = "textBox11";
            this.textBox11.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.textBox11.Size = new System.Drawing.Size(183, 20);
            this.textBox11.TabIndex = 9;
            // 
            // appartmentNum
            // 
            this.appartmentNum.Location = new System.Drawing.Point(108, 125);
            this.appartmentNum.Name = "appartmentNum";
            this.appartmentNum.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.appartmentNum.Size = new System.Drawing.Size(183, 20);
            this.appartmentNum.TabIndex = 8;
            this.appartmentNum.TextChanged += new System.EventHandler(this.appartmentNum_TextChanged);
            // 
            // buildingNum
            // 
            this.buildingNum.Location = new System.Drawing.Point(108, 99);
            this.buildingNum.Name = "buildingNum";
            this.buildingNum.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.buildingNum.Size = new System.Drawing.Size(183, 20);
            this.buildingNum.TabIndex = 7;
            this.buildingNum.TextChanged += new System.EventHandler(this.buildingNum_TextChanged);
            // 
            // town
            // 
            this.town.Location = new System.Drawing.Point(108, 16);
            this.town.Name = "town";
            this.town.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.town.Size = new System.Drawing.Size(183, 20);
            this.town.TabIndex = 5;
            this.town.Text = "г.Владикавказ";
            this.town.TextChanged += new System.EventHandler(this.town_TextChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(5, 151);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(52, 13);
            this.label11.TabIndex = 4;
            this.label11.Text = "Телефон";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(5, 125);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(55, 13);
            this.label12.TabIndex = 3;
            this.label12.Text = "Квартира";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(5, 99);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(43, 13);
            this.label16.TabIndex = 2;
            this.label16.Text = "Корпус";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(5, 42);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(39, 13);
            this.label17.TabIndex = 1;
            this.label17.Text = "Улица";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(5, 16);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(37, 13);
            this.label18.TabIndex = 0;
            this.label18.Text = "Город";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label26);
            this.groupBox3.Controls.Add(this.bankName);
            this.groupBox3.Controls.Add(this.calculationDate);
            this.groupBox3.Controls.Add(this.contractDate);
            this.groupBox3.Controls.Add(this.contractNum);
            this.groupBox3.Controls.Add(this.label13);
            this.groupBox3.Controls.Add(this.label14);
            this.groupBox3.Controls.Add(this.label15);
            this.groupBox3.Location = new System.Drawing.Point(341, 357);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(292, 129);
            this.groupBox3.TabIndex = 10;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Данные договора";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(5, 94);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(32, 13);
            this.label26.TabIndex = 11;
            this.label26.Text = "Банк";
            // 
            // bankName
            // 
            this.bankName.Location = new System.Drawing.Point(108, 94);
            this.bankName.Name = "bankName";
            this.bankName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.bankName.Size = new System.Drawing.Size(110, 20);
            this.bankName.TabIndex = 10;
            // 
            // calculationDate
            // 
            this.calculationDate.CustomFormat = "dd/MM/yy";
            this.calculationDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.calculationDate.Location = new System.Drawing.Point(108, 45);
            this.calculationDate.Name = "calculationDate";
            this.calculationDate.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.calculationDate.Size = new System.Drawing.Size(110, 20);
            this.calculationDate.TabIndex = 8;
            // 
            // contractDate
            // 
            this.contractDate.CustomFormat = "dd/MM/yy";
            this.contractDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.contractDate.Location = new System.Drawing.Point(108, 19);
            this.contractDate.Name = "contractDate";
            this.contractDate.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.contractDate.Size = new System.Drawing.Size(110, 20);
            this.contractDate.TabIndex = 7;
            // 
            // contractNum
            // 
            this.contractNum.Location = new System.Drawing.Point(108, 67);
            this.contractNum.Name = "contractNum";
            this.contractNum.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.contractNum.Size = new System.Drawing.Size(110, 20);
            this.contractNum.TabIndex = 9;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(5, 67);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(91, 13);
            this.label13.TabIndex = 2;
            this.label13.Text = "Номер договора";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(5, 42);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(72, 13);
            this.label14.TabIndex = 1;
            this.label14.Text = "Дата оценки";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(5, 16);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(83, 13);
            this.label15.TabIndex = 0;
            this.label15.Text = "Дата договора";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.showOwnersList);
            this.groupBox2.Controls.Add(this.addOwner);
            this.groupBox2.Controls.Add(this.ownerPassDate);
            this.groupBox2.Controls.Add(this.label32);
            this.groupBox2.Controls.Add(this.ownerPassport);
            this.groupBox2.Controls.Add(this.label33);
            this.groupBox2.Controls.Add(this.label34);
            this.groupBox2.Controls.Add(this.label35);
            this.groupBox2.Controls.Add(this.ownerPassOVD);
            this.groupBox2.Controls.Add(this.ownerPassNum);
            this.groupBox2.Controls.Add(this.button10);
            this.groupBox2.Controls.Add(this.ownerPhone);
            this.groupBox2.Controls.Add(this.ownerAddress);
            this.groupBox2.Controls.Add(this.ownerInit);
            this.groupBox2.Controls.Add(this.ownerName);
            this.groupBox2.Controls.Add(this.ownerSurname);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Location = new System.Drawing.Point(3, 347);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.groupBox2.Size = new System.Drawing.Size(327, 321);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Собственник";
            // 
            // showOwnersList
            // 
            this.showOwnersList.Location = new System.Drawing.Point(8, 284);
            this.showOwnersList.Name = "showOwnersList";
            this.showOwnersList.Size = new System.Drawing.Size(146, 23);
            this.showOwnersList.TabIndex = 29;
            this.showOwnersList.Text = "Список владельцев";
            this.showOwnersList.UseVisualStyleBackColor = true;
            this.showOwnersList.Click += new System.EventHandler(this.showOwnersList_Click);
            // 
            // addOwner
            // 
            this.addOwner.Location = new System.Drawing.Point(175, 255);
            this.addOwner.Name = "addOwner";
            this.addOwner.Size = new System.Drawing.Size(146, 23);
            this.addOwner.TabIndex = 28;
            this.addOwner.Text = "Добавить владельца";
            this.addOwner.UseVisualStyleBackColor = true;
            this.addOwner.Click += new System.EventHandler(this.addOwner_Click);
            // 
            // ownerPassDate
            // 
            this.ownerPassDate.CustomFormat = "dd/MM/yy";
            this.ownerPassDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.ownerPassDate.Location = new System.Drawing.Point(108, 227);
            this.ownerPassDate.Name = "ownerPassDate";
            this.ownerPassDate.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerPassDate.Size = new System.Drawing.Size(204, 20);
            this.ownerPassDate.TabIndex = 13;
            // 
            // label32
            // 
            this.label32.AutoSize = true;
            this.label32.Location = new System.Drawing.Point(5, 145);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(83, 13);
            this.label32.TabIndex = 27;
            this.label32.Text = "Паспорт серия";
            // 
            // ownerPassport
            // 
            this.ownerPassport.Location = new System.Drawing.Point(108, 145);
            this.ownerPassport.Name = "ownerPassport";
            this.ownerPassport.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerPassport.Size = new System.Drawing.Size(204, 20);
            this.ownerPassport.TabIndex = 10;
            // 
            // label33
            // 
            this.label33.AutoSize = true;
            this.label33.Location = new System.Drawing.Point(5, 230);
            this.label33.Name = "label33";
            this.label33.Size = new System.Drawing.Size(73, 13);
            this.label33.TabIndex = 25;
            this.label33.Text = "Дата выдачи";
            // 
            // label34
            // 
            this.label34.AutoSize = true;
            this.label34.Location = new System.Drawing.Point(5, 204);
            this.label34.Name = "label34";
            this.label34.Size = new System.Drawing.Size(40, 13);
            this.label34.TabIndex = 24;
            this.label34.Text = "Выдан";
            // 
            // label35
            // 
            this.label35.AutoSize = true;
            this.label35.Location = new System.Drawing.Point(5, 175);
            this.label35.Name = "label35";
            this.label35.Size = new System.Drawing.Size(18, 13);
            this.label35.TabIndex = 23;
            this.label35.Text = "№";
            // 
            // ownerPassOVD
            // 
            this.ownerPassOVD.Location = new System.Drawing.Point(108, 201);
            this.ownerPassOVD.Name = "ownerPassOVD";
            this.ownerPassOVD.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerPassOVD.Size = new System.Drawing.Size(204, 20);
            this.ownerPassOVD.TabIndex = 12;
            // 
            // ownerPassNum
            // 
            this.ownerPassNum.Location = new System.Drawing.Point(108, 175);
            this.ownerPassNum.Name = "ownerPassNum";
            this.ownerPassNum.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerPassNum.Size = new System.Drawing.Size(204, 20);
            this.ownerPassNum.TabIndex = 11;
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(8, 255);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(141, 23);
            this.button10.TabIndex = 14;
            this.button10.Text = "Проверить падежи";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // ownerPhone
            // 
            this.ownerPhone.Location = new System.Drawing.Point(108, 119);
            this.ownerPhone.Name = "ownerPhone";
            this.ownerPhone.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerPhone.Size = new System.Drawing.Size(204, 20);
            this.ownerPhone.TabIndex = 9;
            // 
            // ownerAddress
            // 
            this.ownerAddress.Location = new System.Drawing.Point(108, 93);
            this.ownerAddress.Name = "ownerAddress";
            this.ownerAddress.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerAddress.Size = new System.Drawing.Size(204, 20);
            this.ownerAddress.TabIndex = 8;
            // 
            // ownerInit
            // 
            this.ownerInit.Location = new System.Drawing.Point(108, 67);
            this.ownerInit.Name = "ownerInit";
            this.ownerInit.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerInit.Size = new System.Drawing.Size(204, 20);
            this.ownerInit.TabIndex = 7;
            this.ownerInit.TextChanged += new System.EventHandler(this.ownerSurname_TextChanged);
            // 
            // ownerName
            // 
            this.ownerName.Location = new System.Drawing.Point(108, 42);
            this.ownerName.Name = "ownerName";
            this.ownerName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerName.Size = new System.Drawing.Size(204, 20);
            this.ownerName.TabIndex = 6;
            this.ownerName.TextChanged += new System.EventHandler(this.ownerSurname_TextChanged);
            // 
            // ownerSurname
            // 
            this.ownerSurname.Location = new System.Drawing.Point(108, 16);
            this.ownerSurname.Name = "ownerSurname";
            this.ownerSurname.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ownerSurname.Size = new System.Drawing.Size(204, 20);
            this.ownerSurname.TabIndex = 5;
            this.ownerSurname.TextChanged += new System.EventHandler(this.ownerSurname_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(5, 119);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 13);
            this.label6.TabIndex = 4;
            this.label6.Text = "Телефон";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(5, 93);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(38, 13);
            this.label7.TabIndex = 3;
            this.label7.Text = "Адрес";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(5, 67);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(54, 13);
            this.label8.TabIndex = 2;
            this.label8.Text = "Отчество";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(5, 42);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(29, 13);
            this.label9.TabIndex = 1;
            this.label9.Text = "Имя";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(5, 16);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(56, 13);
            this.label10.TabIndex = 0;
            this.label10.Text = "Фамилия";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.customerPassDate);
            this.groupBox1.Controls.Add(this.label31);
            this.groupBox1.Controls.Add(this.customerPassport);
            this.groupBox1.Controls.Add(this.label30);
            this.groupBox1.Controls.Add(this.label29);
            this.groupBox1.Controls.Add(this.label23);
            this.groupBox1.Controls.Add(this.customerPassOVD);
            this.groupBox1.Controls.Add(this.customerPassNum);
            this.groupBox1.Controls.Add(this.customerPadBut);
            this.groupBox1.Controls.Add(this.ownerSameCustomer);
            this.groupBox1.Controls.Add(this.customerPhone);
            this.groupBox1.Controls.Add(this.customerAddres);
            this.groupBox1.Controls.Add(this.customerInit);
            this.groupBox1.Controls.Add(this.customerName);
            this.groupBox1.Controls.Add(this.customerSurname);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.groupBox1.Size = new System.Drawing.Size(327, 338);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Заказчик";
            // 
            // customerPassDate
            // 
            this.customerPassDate.CustomFormat = "dd/MM/yy";
            this.customerPassDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.customerPassDate.Location = new System.Drawing.Point(108, 227);
            this.customerPassDate.Name = "customerPassDate";
            this.customerPassDate.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.customerPassDate.Size = new System.Drawing.Size(204, 20);
            this.customerPassDate.TabIndex = 13;
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Location = new System.Drawing.Point(5, 145);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(83, 13);
            this.label31.TabIndex = 19;
            this.label31.Text = "Паспорт серия";
            // 
            // customerPassport
            // 
            this.customerPassport.Location = new System.Drawing.Point(108, 145);
            this.customerPassport.Name = "customerPassport";
            this.customerPassport.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.customerPassport.Size = new System.Drawing.Size(204, 20);
            this.customerPassport.TabIndex = 10;
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.Location = new System.Drawing.Point(5, 230);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(73, 13);
            this.label30.TabIndex = 17;
            this.label30.Text = "Дата выдачи";
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Location = new System.Drawing.Point(5, 204);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(40, 13);
            this.label29.TabIndex = 16;
            this.label29.Text = "Выдан";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(5, 175);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(18, 13);
            this.label23.TabIndex = 15;
            this.label23.Text = "№";
            // 
            // customerPassOVD
            // 
            this.customerPassOVD.Location = new System.Drawing.Point(108, 201);
            this.customerPassOVD.Name = "customerPassOVD";
            this.customerPassOVD.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.customerPassOVD.Size = new System.Drawing.Size(204, 20);
            this.customerPassOVD.TabIndex = 12;
            // 
            // customerPassNum
            // 
            this.customerPassNum.Location = new System.Drawing.Point(108, 175);
            this.customerPassNum.Name = "customerPassNum";
            this.customerPassNum.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.customerPassNum.Size = new System.Drawing.Size(204, 20);
            this.customerPassNum.TabIndex = 11;
            // 
            // customerPadBut
            // 
            this.customerPadBut.Location = new System.Drawing.Point(3, 309);
            this.customerPadBut.Name = "customerPadBut";
            this.customerPadBut.Size = new System.Drawing.Size(146, 23);
            this.customerPadBut.TabIndex = 15;
            this.customerPadBut.Text = "Проверить падежи";
            this.customerPadBut.UseVisualStyleBackColor = true;
            this.customerPadBut.Click += new System.EventHandler(this.customerPadBut_Click);
            // 
            // ownerSameCustomer
            // 
            this.ownerSameCustomer.AutoSize = true;
            this.ownerSameCustomer.Location = new System.Drawing.Point(3, 286);
            this.ownerSameCustomer.Name = "ownerSameCustomer";
            this.ownerSameCustomer.Size = new System.Drawing.Size(205, 17);
            this.ownerSameCustomer.TabIndex = 14;
            this.ownerSameCustomer.Text = "Собственник и заказчик одно лицо";
            this.ownerSameCustomer.UseVisualStyleBackColor = true;
            this.ownerSameCustomer.CheckedChanged += new System.EventHandler(this.ownerSameCustomer_CheckedChanged);
            // 
            // customerPhone
            // 
            this.customerPhone.Location = new System.Drawing.Point(108, 119);
            this.customerPhone.Name = "customerPhone";
            this.customerPhone.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.customerPhone.Size = new System.Drawing.Size(204, 20);
            this.customerPhone.TabIndex = 9;
            // 
            // customerAddres
            // 
            this.customerAddres.Location = new System.Drawing.Point(108, 93);
            this.customerAddres.Name = "customerAddres";
            this.customerAddres.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.customerAddres.Size = new System.Drawing.Size(204, 20);
            this.customerAddres.TabIndex = 8;
            // 
            // customerInit
            // 
            this.customerInit.Location = new System.Drawing.Point(108, 67);
            this.customerInit.Name = "customerInit";
            this.customerInit.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.customerInit.Size = new System.Drawing.Size(204, 20);
            this.customerInit.TabIndex = 7;
            this.customerInit.TextChanged += new System.EventHandler(this.customerSurname_TextChanged);
            // 
            // customerName
            // 
            this.customerName.Location = new System.Drawing.Point(108, 42);
            this.customerName.Name = "customerName";
            this.customerName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.customerName.Size = new System.Drawing.Size(204, 20);
            this.customerName.TabIndex = 6;
            this.customerName.TextChanged += new System.EventHandler(this.customerSurname_TextChanged);
            // 
            // customerSurname
            // 
            this.customerSurname.Location = new System.Drawing.Point(108, 16);
            this.customerSurname.Name = "customerSurname";
            this.customerSurname.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.customerSurname.Size = new System.Drawing.Size(204, 20);
            this.customerSurname.TabIndex = 5;
            this.customerSurname.TextChanged += new System.EventHandler(this.customerSurname_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(5, 119);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(52, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Телефон";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(5, 93);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(38, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Адрес";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(5, 67);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Отчество";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Имя";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Фамилия";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.contractPage);
            this.tabControl1.Controls.Add(this.appartmentsPage);
            this.tabControl1.Controls.Add(this.appartmentsCalcPage);
            this.tabControl1.Controls.Add(this.housePage);
            this.tabControl1.Controls.Add(this.houseCalcPage);
            this.tabControl1.Controls.Add(this.dirtCalcPage);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1142, 742);
            this.tabControl1.TabIndex = 1;
            // 
            // form1BindingSource
            // 
            this.form1BindingSource.DataSource = typeof(WindowsFormsApplication1.mainForm);
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1142, 742);
            this.Controls.Add(this.tabControl1);
            this.Name = "mainForm";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Text = "Оценка";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.mainForm_Load);
            this.dirtCalcPage.ResumeLayout(false);
            this.dirtCalcPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dirtGridAnalogs)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dirtCalcGrid)).EndInit();
            this.houseCalcPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.houseAnalogs)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.houseCalcGrid)).EndInit();
            this.housePage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.appartmentsCalcPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.analogsGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.calculationAppartaments)).EndInit();
            this.appartmentsPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.objectDataGrid)).EndInit();
            this.contractPage.ResumeLayout(false);
            this.contractPage.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.floors)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.roomsNum)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.floor)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.form1BindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.BindingSource form1BindingSource;
        private System.Windows.Forms.TabPage dirtCalcPage;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.DataGridView dirtCalcGrid;
        private System.Windows.Forms.TabPage houseCalcPage;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.DataGridView houseCalcGrid;
        private System.Windows.Forms.TabPage housePage;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.TabPage appartmentsCalcPage;
        private System.Windows.Forms.Button saveAddsAppartaments;
        private System.Windows.Forms.DataGridView analogsGrid;
        private System.Windows.Forms.Button saveResultButton;
        private System.Windows.Forms.Button saveAppartmentsCalc;
        private System.Windows.Forms.DataGridView calculationAppartaments;
        private System.Windows.Forms.TabPage appartmentsPage;
        private System.Windows.Forms.DataGridView objectDataGrid;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TabPage contractPage;
        private System.Windows.Forms.Button loadXMLButton;
        private System.Windows.Forms.TextBox lm2text;
        private System.Windows.Forms.TextBox ownerDocs;
        private System.Windows.Forms.TextBox m2text;
        private System.Windows.Forms.TextBox registrationDoc;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.Label label36;
        private System.Windows.Forms.Label label37;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button gopage2;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox MO;
        private System.Windows.Forms.Label Район;
        private System.Windows.Forms.NumericUpDown floors;
        private System.Windows.Forms.TextBox houseNum;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.NumericUpDown roomsNum;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.ComboBox street;
        private System.Windows.Forms.ComboBox lift;
        private System.Windows.Forms.ComboBox houseType;
        private System.Windows.Forms.NumericUpDown floor;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox textBox11;
        private System.Windows.Forms.TextBox appartmentNum;
        private System.Windows.Forms.TextBox buildingNum;
        private System.Windows.Forms.TextBox town;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.TextBox bankName;
        private System.Windows.Forms.DateTimePicker calculationDate;
        private System.Windows.Forms.DateTimePicker contractDate;
        private System.Windows.Forms.TextBox contractNum;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DateTimePicker ownerPassDate;
        private System.Windows.Forms.Label label32;
        private System.Windows.Forms.TextBox ownerPassport;
        private System.Windows.Forms.Label label33;
        private System.Windows.Forms.Label label34;
        private System.Windows.Forms.Label label35;
        private System.Windows.Forms.TextBox ownerPassOVD;
        private System.Windows.Forms.TextBox ownerPassNum;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.TextBox ownerPhone;
        private System.Windows.Forms.TextBox ownerAddress;
        private System.Windows.Forms.TextBox ownerInit;
        private System.Windows.Forms.TextBox ownerName;
        private System.Windows.Forms.TextBox ownerSurname;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker customerPassDate;
        private System.Windows.Forms.Label label31;
        private System.Windows.Forms.TextBox customerPassport;
        private System.Windows.Forms.Label label30;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.TextBox customerPassOVD;
        private System.Windows.Forms.TextBox customerPassNum;
        private System.Windows.Forms.Button customerPadBut;
        private System.Windows.Forms.CheckBox ownerSameCustomer;
        private System.Windows.Forms.TextBox customerPhone;
        private System.Windows.Forms.TextBox customerAddres;
        private System.Windows.Forms.TextBox customerInit;
        private System.Windows.Forms.TextBox customerName;
        private System.Windows.Forms.TextBox customerSurname;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.TextBox orgAdd;
        private System.Windows.Forms.Label label38;
        private System.Windows.Forms.TextBox orgOGRN;
        private System.Windows.Forms.TextBox orgINN;
        private System.Windows.Forms.TextBox orgName;
        private System.Windows.Forms.Label label42;
        private System.Windows.Forms.Label label43;
        private System.Windows.Forms.Label label44;
        private System.Windows.Forms.Label label45;
        private System.Windows.Forms.Label label46;
        private System.Windows.Forms.CheckBox ownerOrg;
        private System.Windows.Forms.DateTimePicker orgRegDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn DirtName;
        private System.Windows.Forms.DataGridViewTextBoxColumn A1;
        private System.Windows.Forms.DataGridViewTextBoxColumn A2;
        private System.Windows.Forms.DataGridViewTextBoxColumn A3;
        private System.Windows.Forms.CheckBox newBuildingCheck;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label39;
        private System.Windows.Forms.Button addOwner;
        private System.Windows.Forms.Button SaveXMLButton;
        private System.Windows.Forms.Button showOwnersList;
        private System.Windows.Forms.DataGridView houseAnalogs;
        private System.Windows.Forms.DataGridView dirtGridAnalogs;
        private System.Windows.Forms.Button saveGridToWordButton;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label41;
        private System.Windows.Forms.TextBox dirtKadastr;
        private System.Windows.Forms.Label label40;
        private System.Windows.Forms.TextBox dirtm2;
        private System.Windows.Forms.TextBox gridDoc2;
        private System.Windows.Forms.TextBox gridDoc;
        private System.Windows.Forms.Label label47;
        private System.Windows.Forms.Label label48;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.TextBox orgKPP;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn18;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn19;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn20;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn21;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn13;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn14;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn15;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn16;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn17;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn11;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn12;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column7;

        public mainForm(string type, string banks)
        {
            InitializeComponent();
            Missing = System.Reflection.Missing.Value;
            int t = tabControl1.TabPages.Count;
            for (int i = 0; i < t; i++)
            {
                //tabControl1.TabPages[i].;            
            }
            docTypeT = type;
            switch (type)
            {
                case "Квартира":
                    {
                        string fileName = System.Windows.Forms.Application.StartupPath + "\\calcState.xml";
                        
                        tabControl1.TabPages.Remove(tabControl1.TabPages[3]);
                        tabControl1.TabPages.Remove(tabControl1.TabPages[3]);
                        tabControl1.TabPages.Remove(tabControl1.TabPages[3]);
                        //tabControl1.TabPages.Remove(tabControl1.TabPages[3]);
                       

                        addObjectData();
                        calculationAppartaments.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
                        calculationAppartaments.AutoResizeRows();
                        calculationAppartaments.AutoResizeColumns();
                        //analogsGrid.AutoSizeRowsMode  = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
                        //analogsGrid.AutoResizeRows();
                        //analogsGrid.AutoResizeColumns();
                        docType = type.ToLower();
                        System.Data.DataTable test = getDataFromXLS("Черновик.xls");
                        calculationAppartaments.DataSource = test;
                        calculationAppartaments.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                        calculationAppartaments.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                        calculationAppartaments.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                        calculationAppartaments.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                        calculationAppartaments.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
                        //calculateCost();
                        test = null;
                        test = getDataFromXLS("analogs.xls");
                        analogsGrid.DataSource = test;
                        analogsGrid.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                        analogsGrid.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                        analogsGrid.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                        analogsGrid.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                        loadState(fileName);

                    } break;
                case "Домовладение":
                    {
                        tabControl1.TabPages.Remove(tabControl1.TabPages[1]);
                        tabControl1.TabPages.Remove(tabControl1.TabPages[1]);
                        tabControl1.TabPages.Remove(tabControl1.TabPages[3]);
                        //tabControl1.TabPages.Remove(tabControl1.TabPages[3]);
                        addHouseData();
                        docType = type.ToLower();
                        houseCalcGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
                        houseCalcGrid.AutoResizeRows();
                        houseCalcGrid.AutoResizeColumns();
                        System.Data.DataTable test = getDataFromXLS("analogsHouse.xls");

                        houseAnalogs.DataSource = test;
                        houseAnalogs.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                        houseAnalogs.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                        houseAnalogs.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                        houseAnalogs.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;

                    }
                    break;
                case "Земельный участок":
                    {
                        tabControl1.TabPages.Remove(tabControl1.TabPages[1]);
                        tabControl1.TabPages.Remove(tabControl1.TabPages[1]);
                        tabControl1.TabPages.Remove(tabControl1.TabPages[1]);
                        tabControl1.TabPages.Remove(tabControl1.TabPages[1]);
                        addGridData();
                        docType = type.ToLower();
                        dirtCalcGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
                        dirtCalcGrid.AutoResizeRows();
                        dirtCalcGrid.AutoResizeColumns();
                        System.Data.DataTable test = getDataFromXLS("analogsDirt.xls");
                        saveGridToWordButton.Show();
                        dirtGridAnalogs.DataSource = test;
                       
                    }
                    break;
                case "Домовладение с земельным участком":
                    {
                        tabControl1.TabPages.Remove(tabControl1.TabPages[1]);
                        tabControl1.TabPages.Remove(tabControl1.TabPages[1]);
                        addGridData();
                        addHouseData();
                        houseCalcGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
                        houseCalcGrid.AutoResizeRows();
                        houseCalcGrid.AutoResizeColumns();
                        System.Data.DataTable test = getDataFromXLS("дом.xls");
                        houseCalcGrid.DataSource = test;
                        test = null;
                        test = getDataFromXLS("analogsHouse.xls");
                        houseAnalogs.DataSource = test;
                        houseAnalogs.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                        houseAnalogs.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                        houseAnalogs.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                        houseAnalogs.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dirtCalcGrid.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dirtCalcGrid.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dirtCalcGrid.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dirtCalcGrid.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                        dirtCalcGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
                        dirtCalcGrid.AutoResizeRows();
                        dirtCalcGrid.AutoResizeColumns();
                        docType = type.ToLower();
                        test = null;
                        test = getDataFromXLS("analogsDirt.xls");

                        dirtGridAnalogs.DataSource = test;
                        
                        ////dirtCalcGrid.Columns.RemoveAt(0);
                        ////dirtCalcGrid.Columns.RemoveAt(0);
                        ////dirtCalcGrid.Columns.RemoveAt(0);
                        ////dirtCalcGrid.Columns.RemoveAt(0);

                        dirtGridAnalogs.Columns.RemoveAt(0);
                        dirtGridAnalogs.Columns.RemoveAt(0);
                        dirtGridAnalogs.Columns.RemoveAt(0);
                        dirtGridAnalogs.Columns.RemoveAt(0);

                        houseAnalogs.Columns.RemoveAt(0);
                        houseAnalogs.Columns.RemoveAt(0);
                        houseAnalogs.Columns.RemoveAt(0);
                        houseAnalogs.Columns.RemoveAt(0);
                        houseAnalogs.Columns.RemoveAt(0);

                        houseCalcGrid.Columns.RemoveAt(0);
                        houseCalcGrid.Columns.RemoveAt(0);
                        houseCalcGrid.Columns.RemoveAt(0);
                        houseCalcGrid.Columns.RemoveAt(0);
                        houseCalcGrid.Columns.RemoveAt(0);
                    }
                    break;


                default: break;
            }

            Bank bank = new Bank();
            bank.BankName = banks;
            bankName.Text = bank.BankName;
            

        }
    }
}

