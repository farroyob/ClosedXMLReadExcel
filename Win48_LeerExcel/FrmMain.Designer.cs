namespace Win48_LeerExcel
{
    partial class FrmMain
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
            this.txtDir = new System.Windows.Forms.TextBox();
            this.btnRead = new System.Windows.Forms.Button();
            this.lblFindFile = new System.Windows.Forms.Label();
            this.dtgData = new System.Windows.Forms.DataGridView();
            this.gpbMain = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.dtgData)).BeginInit();
            this.gpbMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtDir
            // 
            this.txtDir.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDir.Location = new System.Drawing.Point(43, 16);
            this.txtDir.Name = "txtDir";
            this.txtDir.Size = new System.Drawing.Size(593, 20);
            this.txtDir.TabIndex = 0;
            // 
            // btnRead
            // 
            this.btnRead.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRead.Location = new System.Drawing.Point(641, 15);
            this.btnRead.Name = "btnRead";
            this.btnRead.Size = new System.Drawing.Size(75, 23);
            this.btnRead.TabIndex = 1;
            this.btnRead.Text = "Read";
            this.btnRead.UseVisualStyleBackColor = true;
            this.btnRead.Click += new System.EventHandler(this.btnRead_Click);
            // 
            // lblFindFile
            // 
            this.lblFindFile.AutoSize = true;
            this.lblFindFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFindFile.Location = new System.Drawing.Point(6, 20);
            this.lblFindFile.Name = "lblFindFile";
            this.lblFindFile.Size = new System.Drawing.Size(31, 13);
            this.lblFindFile.TabIndex = 2;
            this.lblFindFile.Text = "File:";
            this.lblFindFile.Click += new System.EventHandler(this.lblFindFile_Click);
            // 
            // dtgData
            // 
            this.dtgData.AllowUserToAddRows = false;
            this.dtgData.AllowUserToDeleteRows = false;
            this.dtgData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtgData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtgData.Location = new System.Drawing.Point(17, 66);
            this.dtgData.Name = "dtgData";
            this.dtgData.ReadOnly = true;
            this.dtgData.Size = new System.Drawing.Size(722, 359);
            this.dtgData.TabIndex = 3;
            // 
            // gpbMain
            // 
            this.gpbMain.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gpbMain.Controls.Add(this.lblFindFile);
            this.gpbMain.Controls.Add(this.txtDir);
            this.gpbMain.Controls.Add(this.btnRead);
            this.gpbMain.Location = new System.Drawing.Point(17, 12);
            this.gpbMain.Name = "gpbMain";
            this.gpbMain.Size = new System.Drawing.Size(722, 48);
            this.gpbMain.TabIndex = 4;
            this.gpbMain.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(756, 437);
            this.Controls.Add(this.gpbMain);
            this.Controls.Add(this.dtgData);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Main";
            ((System.ComponentModel.ISupportInitialize)(this.dtgData)).EndInit();
            this.gpbMain.ResumeLayout(false);
            this.gpbMain.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox txtDir;
        private System.Windows.Forms.Button btnRead;
        private System.Windows.Forms.Label lblFindFile;
        private System.Windows.Forms.DataGridView dtgData;
        private System.Windows.Forms.GroupBox gpbMain;
    }
}

