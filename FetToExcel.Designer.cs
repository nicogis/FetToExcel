namespace FetToExcel
{
    partial class FetToExcel
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnFileFet = new System.Windows.Forms.Button();
            this.txtFileFet = new System.Windows.Forms.TextBox();
            this.txtFileExcel = new System.Windows.Forms.TextBox();
            this.btnFileExcel = new System.Windows.Forms.Button();
            this.btnImporta = new System.Windows.Forms.Button();
            this.saveFD = new System.Windows.Forms.SaveFileDialog();
            this.openFD = new System.Windows.Forms.OpenFileDialog();
            this.llinkFetToExcel = new System.Windows.Forms.LinkLabel();
            this.txtStartCell = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnFileFet
            // 
            this.btnFileFet.Location = new System.Drawing.Point(393, 21);
            this.btnFileFet.Name = "btnFileFet";
            this.btnFileFet.Size = new System.Drawing.Size(104, 23);
            this.btnFileFet.TabIndex = 0;
            this.btnFileFet.Text = "File teachers.xml";
            this.btnFileFet.UseVisualStyleBackColor = true;
            this.btnFileFet.Click += new System.EventHandler(this.BtnFileFet_Click);
            // 
            // txtFileFet
            // 
            this.txtFileFet.Location = new System.Drawing.Point(17, 21);
            this.txtFileFet.Name = "txtFileFet";
            this.txtFileFet.ReadOnly = true;
            this.txtFileFet.Size = new System.Drawing.Size(370, 20);
            this.txtFileFet.TabIndex = 1;
            // 
            // txtFileExcel
            // 
            this.txtFileExcel.Location = new System.Drawing.Point(17, 55);
            this.txtFileExcel.Name = "txtFileExcel";
            this.txtFileExcel.ReadOnly = true;
            this.txtFileExcel.Size = new System.Drawing.Size(370, 20);
            this.txtFileExcel.TabIndex = 3;
            // 
            // btnFileExcel
            // 
            this.btnFileExcel.Location = new System.Drawing.Point(393, 55);
            this.btnFileExcel.Name = "btnFileExcel";
            this.btnFileExcel.Size = new System.Drawing.Size(104, 23);
            this.btnFileExcel.TabIndex = 2;
            this.btnFileExcel.Text = "File output Excel";
            this.btnFileExcel.UseVisualStyleBackColor = true;
            this.btnFileExcel.Click += new System.EventHandler(this.BtnFileExcel_Click);
            // 
            // btnImporta
            // 
            this.btnImporta.Location = new System.Drawing.Point(216, 118);
            this.btnImporta.Name = "btnImporta";
            this.btnImporta.Size = new System.Drawing.Size(75, 23);
            this.btnImporta.TabIndex = 4;
            this.btnImporta.Text = "Genera";
            this.btnImporta.UseVisualStyleBackColor = true;
            this.btnImporta.Click += new System.EventHandler(this.BtnImporta_Click);
            // 
            // saveFD
            // 
            this.saveFD.DefaultExt = "xlsx";
            this.saveFD.Filter = "File Excel|*xlsx";
            this.saveFD.FileOk += new System.ComponentModel.CancelEventHandler(this.SaveFD_FileOk);
            // 
            // openFD
            // 
            this.openFD.DefaultExt = "xml";
            this.openFD.Filter = "File fet|*.xml";
            // 
            // llinkFetToExcel
            // 
            this.llinkFetToExcel.AutoSize = true;
            this.llinkFetToExcel.Location = new System.Drawing.Point(405, 128);
            this.llinkFetToExcel.Name = "llinkFetToExcel";
            this.llinkFetToExcel.Size = new System.Drawing.Size(63, 13);
            this.llinkFetToExcel.TabIndex = 5;
            this.llinkFetToExcel.TabStop = true;
            this.llinkFetToExcel.Text = "Fet to Excel";
            this.llinkFetToExcel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LlinkFetToExcel_LinkClicked);
            // 
            // txtStartCell
            // 
            this.txtStartCell.Location = new System.Drawing.Point(108, 91);
            this.txtStartCell.Name = "txtStartCell";
            this.txtStartCell.Size = new System.Drawing.Size(40, 20);
            this.txtStartCell.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(14, 94);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Cella di partenza:";
            // 
            // FetToExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(506, 153);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtStartCell);
            this.Controls.Add(this.llinkFetToExcel);
            this.Controls.Add(this.btnImporta);
            this.Controls.Add(this.txtFileExcel);
            this.Controls.Add(this.btnFileExcel);
            this.Controls.Add(this.txtFileFet);
            this.Controls.Add(this.btnFileFet);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FetToExcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Fet to Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnFileFet;
        private System.Windows.Forms.TextBox txtFileFet;
        private System.Windows.Forms.TextBox txtFileExcel;
        private System.Windows.Forms.Button btnFileExcel;
        private System.Windows.Forms.Button btnImporta;
        private System.Windows.Forms.SaveFileDialog saveFD;
        private System.Windows.Forms.OpenFileDialog openFD;
        private System.Windows.Forms.LinkLabel llinkFetToExcel;
        private System.Windows.Forms.TextBox txtStartCell;
        private System.Windows.Forms.Label label1;
    }
}

