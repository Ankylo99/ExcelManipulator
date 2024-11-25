namespace ExcelManipulator
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnProcessFile = new System.Windows.Forms.Button();
            this.dtpStartDate = new System.Windows.Forms.DateTimePicker();
            this.dtpEndDate = new System.Windows.Forms.DateTimePicker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F);
            this.label1.Location = new System.Drawing.Point(67, 59);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(127, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "Archivo Excel";
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(223, 61);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(176, 20);
            this.txtFilePath.TabIndex = 1;
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(419, 61);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(112, 23);
            this.btnSelectFile.TabIndex = 2;
            this.btnSelectFile.Text = "Seleccionar Archivo";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F);
            this.label2.Location = new System.Drawing.Point(67, 113);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 24);
            this.label2.TabIndex = 4;
            this.label2.Text = "Fecha Inicio";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F);
            this.label3.Location = new System.Drawing.Point(67, 154);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(96, 24);
            this.label3.TabIndex = 6;
            this.label3.Text = "Fecha Fin";
            // 
            // btnProcessFile
            // 
            this.btnProcessFile.Location = new System.Drawing.Point(488, 265);
            this.btnProcessFile.Name = "btnProcessFile";
            this.btnProcessFile.Size = new System.Drawing.Size(100, 23);
            this.btnProcessFile.TabIndex = 8;
            this.btnProcessFile.Text = "Procesar Archivo";
            this.btnProcessFile.UseVisualStyleBackColor = true;
            this.btnProcessFile.Click += new System.EventHandler(this.btnProcessFile_ClickAsync);
            // 
            // dtpStartDate
            // 
            this.dtpStartDate.Location = new System.Drawing.Point(223, 113);
            this.dtpStartDate.Name = "dtpStartDate";
            this.dtpStartDate.Size = new System.Drawing.Size(200, 20);
            this.dtpStartDate.TabIndex = 11;
            // 
            // dtpEndDate
            // 
            this.dtpEndDate.Location = new System.Drawing.Point(223, 154);
            this.dtpEndDate.Name = "dtpEndDate";
            this.dtpEndDate.Size = new System.Drawing.Size(200, 20);
            this.dtpEndDate.TabIndex = 12;
            // 
            // progressBar1
            // 
            this.progressBar1.Visible = false; // Al principio no visible
            this.progressBar1.Location = new System.Drawing.Point(223, 213);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(200, 23);
            this.progressBar1.TabIndex = 13;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // lblStatus
            // 
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.lblStatus.Location = new System.Drawing.Point(223, 240);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(100, 17);
            this.lblStatus.TabIndex = 14;
            this.lblStatus.Text = "";  // Aquí el mensaje inicial está vacío.
            this.lblStatus.Visible = false; // Al principio no visible
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 300);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dtpEndDate);
            this.Controls.Add(this.dtpStartDate);
            this.Controls.Add(this.btnProcessFile);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Manipulador de Excel";
            ((System.ComponentModel.ISupportInitialize)(this.loadingIcon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnProcessFile;
        private System.Windows.Forms.DateTimePicker dtpStartDate;
        private System.Windows.Forms.DateTimePicker dtpEndDate;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label lblStatus;
    }
}
