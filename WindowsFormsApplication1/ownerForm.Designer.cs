namespace HouseCostCalculation
{
    partial class ownerForm
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
            this.ownerList = new System.Windows.Forms.DataGridView();
            this.ownerName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ownerSurname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ownerLastname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ownerAddres = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ownerPassport = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ownerPassNum = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ownerPassOVD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ownerDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.ownerList)).BeginInit();
            this.SuspendLayout();
            // 
            // ownerList
            // 
            this.ownerList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ownerList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ownerName,
            this.ownerSurname,
            this.ownerLastname,
            this.ownerAddres,
            this.ownerPassport,
            this.ownerPassNum,
            this.ownerPassOVD,
            this.ownerDate});
            this.ownerList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ownerList.Location = new System.Drawing.Point(0, 0);
            this.ownerList.Name = "ownerList";
            this.ownerList.Size = new System.Drawing.Size(843, 296);
            this.ownerList.TabIndex = 0;
            // 
            // ownerName
            // 
            this.ownerName.HeaderText = "Имя";
            this.ownerName.Name = "ownerName";
            // 
            // ownerSurname
            // 
            this.ownerSurname.HeaderText = "Отчество";
            this.ownerSurname.Name = "ownerSurname";
            // 
            // ownerLastname
            // 
            this.ownerLastname.HeaderText = "Фамилия";
            this.ownerLastname.Name = "ownerLastname";
            // 
            // ownerAddres
            // 
            this.ownerAddres.HeaderText = "Адрес";
            this.ownerAddres.Name = "ownerAddres";
            // 
            // ownerPassport
            // 
            this.ownerPassport.HeaderText = "Серия паспорта";
            this.ownerPassport.Name = "ownerPassport";
            // 
            // ownerPassNum
            // 
            this.ownerPassNum.HeaderText = "Номер паспорта";
            this.ownerPassNum.Name = "ownerPassNum";
            // 
            // ownerPassOVD
            // 
            this.ownerPassOVD.HeaderText = "Выдан";
            this.ownerPassOVD.Name = "ownerPassOVD";
            // 
            // ownerDate
            // 
            this.ownerDate.HeaderText = "Дата выдачи";
            this.ownerDate.Name = "ownerDate";
            // 
            // ownerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(843, 296);
            this.Controls.Add(this.ownerList);
            this.Name = "ownerForm";
            this.Text = "Список владельцев";
            ((System.ComponentModel.ISupportInitialize)(this.ownerList)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView ownerList;
        private System.Windows.Forms.DataGridViewTextBoxColumn ownerName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ownerSurname;
        private System.Windows.Forms.DataGridViewTextBoxColumn ownerLastname;
        private System.Windows.Forms.DataGridViewTextBoxColumn ownerAddres;
        private System.Windows.Forms.DataGridViewTextBoxColumn ownerPassport;
        private System.Windows.Forms.DataGridViewTextBoxColumn ownerPassNum;
        private System.Windows.Forms.DataGridViewTextBoxColumn ownerPassOVD;
        private System.Windows.Forms.DataGridViewTextBoxColumn ownerDate;
    }
}