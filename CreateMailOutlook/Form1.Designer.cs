namespace CreateMailOutlook
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnStart = new Button();
            label1 = new Label();
            txtProxy = new TextBox();
            txtProductId = new TextBox();
            label2 = new Label();
            txtInstallDate = new TextBox();
            label3 = new Label();
            txtProductName = new TextBox();
            label4 = new Label();
            txtMachineId = new TextBox();
            label5 = new Label();
            txtMAC = new TextBox();
            label6 = new Label();
            dataGridView1 = new DataGridView();
            txtLog = new RichTextBox();
            btnStop = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // btnStart
            // 
            btnStart.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnStart.Location = new Point(502, 597);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(94, 29);
            btnStart.TabIndex = 0;
            btnStart.Text = "Start";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnStart_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(24, 23);
            label1.Name = "label1";
            label1.Size = new Size(90, 20);
            label1.TabIndex = 1;
            label1.Text = "Proxy Server";
            // 
            // txtProxy
            // 
            txtProxy.Font = new Font("Segoe UI", 10.2F);
            txtProxy.Location = new Point(129, 20);
            txtProxy.Name = "txtProxy";
            txtProxy.Size = new Size(268, 30);
            txtProxy.TabIndex = 2;
            // 
            // txtProductId
            // 
            txtProductId.Font = new Font("Segoe UI", 10.2F);
            txtProductId.Location = new Point(129, 104);
            txtProductId.Name = "txtProductId";
            txtProductId.Size = new Size(268, 30);
            txtProductId.TabIndex = 4;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(24, 107);
            label2.Name = "label2";
            label2.Size = new Size(73, 20);
            label2.TabIndex = 3;
            label2.Text = "ProductId";
            // 
            // txtInstallDate
            // 
            txtInstallDate.Font = new Font("Segoe UI", 10.2F);
            txtInstallDate.Location = new Point(129, 181);
            txtInstallDate.Name = "txtInstallDate";
            txtInstallDate.Size = new Size(268, 30);
            txtInstallDate.TabIndex = 6;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(24, 184);
            label3.Name = "label3";
            label3.Size = new Size(80, 20);
            label3.TabIndex = 5;
            label3.Text = "InstallDate";
            // 
            // txtProductName
            // 
            txtProductName.Font = new Font("Segoe UI", 10.2F);
            txtProductName.Location = new Point(129, 261);
            txtProductName.Name = "txtProductName";
            txtProductName.Size = new Size(268, 30);
            txtProductName.TabIndex = 8;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(24, 264);
            label4.Name = "label4";
            label4.Size = new Size(100, 20);
            label4.TabIndex = 7;
            label4.Text = "ProductName";
            // 
            // txtMachineId
            // 
            txtMachineId.Font = new Font("Segoe UI", 10.2F);
            txtMachineId.Location = new Point(129, 339);
            txtMachineId.Name = "txtMachineId";
            txtMachineId.Size = new Size(268, 30);
            txtMachineId.TabIndex = 10;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(24, 342);
            label5.Name = "label5";
            label5.Size = new Size(78, 20);
            label5.TabIndex = 9;
            label5.Text = "MachineId";
            // 
            // txtMAC
            // 
            txtMAC.Font = new Font("Segoe UI", 10.2F);
            txtMAC.Location = new Point(129, 436);
            txtMAC.Name = "txtMAC";
            txtMAC.Size = new Size(268, 30);
            txtMAC.TabIndex = 12;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(24, 439);
            label6.Name = "label6";
            label6.Size = new Size(41, 20);
            label6.TabIndex = 11;
            label6.Text = "MAC";
            // 
            // dataGridView1
            // 
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(408, 20);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.Size = new Size(435, 260);
            dataGridView1.TabIndex = 13;
            // 
            // txtLog
            // 
            txtLog.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            txtLog.BackColor = Color.White;
            txtLog.HideSelection = false;
            txtLog.Location = new Point(408, 286);
            txtLog.Name = "txtLog";
            txtLog.Size = new Size(435, 305);
            txtLog.TabIndex = 14;
            txtLog.Text = "";
            // 
            // btnStop
            // 
            btnStop.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnStop.BackColor = Color.Red;
            btnStop.Enabled = false;
            btnStop.Location = new Point(646, 597);
            btnStop.Name = "btnStop";
            btnStop.Size = new Size(94, 29);
            btnStop.TabIndex = 15;
            btnStop.Text = "Stop";
            btnStop.UseVisualStyleBackColor = false;
            btnStop.Click += btnStop_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            ClientSize = new Size(855, 638);
            Controls.Add(btnStop);
            Controls.Add(txtLog);
            Controls.Add(dataGridView1);
            Controls.Add(txtMAC);
            Controls.Add(label6);
            Controls.Add(txtMachineId);
            Controls.Add(label5);
            Controls.Add(txtProductName);
            Controls.Add(label4);
            Controls.Add(txtInstallDate);
            Controls.Add(label3);
            Controls.Add(txtProductId);
            Controls.Add(label2);
            Controls.Add(txtProxy);
            Controls.Add(label1);
            Controls.Add(btnStart);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnStart;
        private Label label1;
        private TextBox txtProxy;
        private TextBox txtProductId;
        private Label label2;
        private TextBox txtInstallDate;
        private Label label3;
        private TextBox txtProductName;
        private Label label4;
        private TextBox txtMachineId;
        private Label label5;
        private TextBox txtMAC;
        private Label label6;
        private DataGridView dataGridView1;
        private RichTextBox txtLog;
        private Button btnStop;
    }
}
