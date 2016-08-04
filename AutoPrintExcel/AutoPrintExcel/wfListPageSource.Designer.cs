namespace AutoPrintExcel
{
    partial class wfListPageSource
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.btnGetPageSource = new System.Windows.Forms.Button();
            this.btnChoosePageSource = new System.Windows.Forms.Button();
            this.panel5 = new System.Windows.Forms.Panel();
            this.groupPageSource = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel5.SuspendLayout();
            this.groupPageSource.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 289);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(672, 40);
            this.panel1.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(672, 40);
            this.panel2.TabIndex = 1;
            // 
            // panel3
            // 
            this.panel3.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel3.Location = new System.Drawing.Point(0, 40);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(50, 249);
            this.panel3.TabIndex = 2;
            // 
            // panel4
            // 
            this.panel4.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel4.Location = new System.Drawing.Point(622, 40);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(50, 249);
            this.panel4.TabIndex = 3;
            // 
            // btnGetPageSource
            // 
            this.btnGetPageSource.Location = new System.Drawing.Point(490, 59);
            this.btnGetPageSource.Name = "btnGetPageSource";
            this.btnGetPageSource.Size = new System.Drawing.Size(126, 49);
            this.btnGetPageSource.TabIndex = 5;
            this.btnGetPageSource.Text = "Get Page Source";
            this.btnGetPageSource.UseVisualStyleBackColor = true;
            this.btnGetPageSource.Click += new System.EventHandler(this.btnGetPageSource_Click);
            // 
            // btnChoosePageSource
            // 
            this.btnChoosePageSource.Location = new System.Drawing.Point(490, 123);
            this.btnChoosePageSource.Name = "btnChoosePageSource";
            this.btnChoosePageSource.Size = new System.Drawing.Size(126, 49);
            this.btnChoosePageSource.TabIndex = 6;
            this.btnChoosePageSource.Text = "Use Page Source";
            this.btnChoosePageSource.UseVisualStyleBackColor = true;
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.groupPageSource);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel5.Location = new System.Drawing.Point(50, 40);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(426, 249);
            this.panel5.TabIndex = 7;
            // 
            // groupPageSource
            // 
            this.groupPageSource.Controls.Add(this.dataGridView1);
            this.groupPageSource.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupPageSource.Location = new System.Drawing.Point(0, 0);
            this.groupPageSource.Name = "groupPageSource";
            this.groupPageSource.Size = new System.Drawing.Size(426, 249);
            this.groupPageSource.TabIndex = 0;
            this.groupPageSource.TabStop = false;
            this.groupPageSource.Text = "Group Page Source";
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(3, 19);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(420, 227);
            this.dataGridView1.TabIndex = 1;
            // 
            // wfListPageSource
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(672, 329);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.btnChoosePageSource);
            this.Controls.Add(this.btnGetPageSource);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "wfListPageSource";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.panel5.ResumeLayout(false);
            this.groupPageSource.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Button btnGetPageSource;
        private System.Windows.Forms.Button btnChoosePageSource;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.GroupBox groupPageSource;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}