
namespace AutoLockWindows
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.btnSave = new System.Windows.Forms.Button();
            this.cbWord = new System.Windows.Forms.CheckBox();
            this.cbExcel = new System.Windows.Forms.CheckBox();
            this.btStop = new System.Windows.Forms.Button();
            this.btRun = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.pnlLogin = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.lblTimeLeft = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.pnlLogin.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 65);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 21);
            this.label1.TabIndex = 0;
            this.label1.Text = "דקות עד נעילה";
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Font = new System.Drawing.Font("Segoe UI", 25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.numericUpDown1.ForeColor = System.Drawing.Color.Maroon;
            this.numericUpDown1.Location = new System.Drawing.Point(25, 102);
            this.numericUpDown1.Margin = new System.Windows.Forms.Padding(4);
            this.numericUpDown1.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.numericUpDown1.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(154, 52);
            this.numericUpDown1.TabIndex = 1;
            this.numericUpDown1.Value = new decimal(new int[] {
            15,
            0,
            0,
            0});
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnSave.ForeColor = System.Drawing.Color.RoyalBlue;
            this.btnSave.Location = new System.Drawing.Point(25, 305);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(154, 35);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "שמור הגדרות";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // cbWord
            // 
            this.cbWord.AutoSize = true;
            this.cbWord.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.cbWord.Location = new System.Drawing.Point(25, 183);
            this.cbWord.Name = "cbWord";
            this.cbWord.Size = new System.Drawing.Size(279, 34);
            this.cbWord.TabIndex = 3;
            this.cbWord.Text = "לא לנעול אם WORD פתוח";
            this.cbWord.UseVisualStyleBackColor = true;
            // 
            // cbExcel
            // 
            this.cbExcel.AutoSize = true;
            this.cbExcel.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.cbExcel.Location = new System.Drawing.Point(25, 238);
            this.cbExcel.Name = "cbExcel";
            this.cbExcel.Size = new System.Drawing.Size(272, 34);
            this.cbExcel.TabIndex = 4;
            this.cbExcel.Text = "לא לנעול אם EXCEL פתוח";
            this.cbExcel.UseVisualStyleBackColor = true;
            // 
            // btStop
            // 
            this.btStop.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btStop.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btStop.ForeColor = System.Drawing.Color.Brown;
            this.btStop.Location = new System.Drawing.Point(364, 308);
            this.btStop.Name = "btStop";
            this.btStop.Size = new System.Drawing.Size(169, 32);
            this.btStop.TabIndex = 5;
            this.btStop.Text = "הפסקת הסטופר";
            this.btStop.UseVisualStyleBackColor = true;
            this.btStop.Click += new System.EventHandler(this.btStop_Click);
            // 
            // btRun
            // 
            this.btRun.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btRun.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btRun.ForeColor = System.Drawing.Color.Green;
            this.btRun.Location = new System.Drawing.Point(364, 275);
            this.btRun.Name = "btRun";
            this.btRun.Size = new System.Drawing.Size(169, 31);
            this.btRun.TabIndex = 6;
            this.btRun.Text = "הפעלת הסטופר";
            this.btRun.UseVisualStyleBackColor = true;
            this.btRun.Click += new System.EventHandler(this.btRun_Click);
            // 
            // listView1
            // 
            this.listView1.BackColor = System.Drawing.Color.Black;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.listView1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.listView1.Font = new System.Drawing.Font("Courier New", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.listView1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(0, 365);
            this.listView1.Name = "listView1";
            this.listView1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.listView1.Size = new System.Drawing.Size(545, 153);
            this.listView1.TabIndex = 8;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Width = 1000;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "notifyIcon1";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.Click += new System.EventHandler(this.notifyIcon1_Click);
            // 
            // linkLabel1
            // 
            this.linkLabel1.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.linkLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline;
            this.linkLabel1.LinkColor = System.Drawing.Color.Blue;
            this.linkLabel1.Location = new System.Drawing.Point(238, 337);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(68, 15);
            this.linkLabel1.TabIndex = 9;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "לנעול עכשיו";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.button1.ForeColor = System.Drawing.Color.Firebrick;
            this.button1.Location = new System.Drawing.Point(471, 48);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(62, 38);
            this.button1.TabIndex = 2;
            this.button1.Text = "צא";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.button2.ForeColor = System.Drawing.Color.MediumPurple;
            this.button2.Location = new System.Drawing.Point(403, 48);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(62, 38);
            this.button2.TabIndex = 2;
            this.button2.Text = "הסתר";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // pnlLogin
            // 
            this.pnlLogin.BackColor = System.Drawing.Color.DimGray;
            this.pnlLogin.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlLogin.Controls.Add(this.pictureBox1);
            this.pnlLogin.Controls.Add(this.button3);
            this.pnlLogin.Controls.Add(this.label2);
            this.pnlLogin.Controls.Add(this.txtPassword);
            this.pnlLogin.Location = new System.Drawing.Point(508, 0);
            this.pnlLogin.Name = "pnlLogin";
            this.pnlLogin.Size = new System.Drawing.Size(36, 26);
            this.pnlLogin.TabIndex = 10;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::AutoLockWindows.Properties.Resources.icon;
            this.pictureBox1.Location = new System.Drawing.Point(256, 118);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(30, 35);
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(90)))), ((int)(((byte)(90)))), ((int)(((byte)(90)))));
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button3.ForeColor = System.Drawing.Color.White;
            this.button3.Location = new System.Drawing.Point(206, 354);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(115, 47);
            this.button3.TabIndex = 2;
            this.button3.Text = "המשך";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Segoe UI", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(46, 216);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(450, 41);
            this.label2.TabIndex = 1;
            this.label2.Text = "סיסמת חשבון משתמש";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // txtPassword
            // 
            this.txtPassword.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.txtPassword.ForeColor = System.Drawing.Color.Black;
            this.txtPassword.Location = new System.Drawing.Point(46, 262);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(450, 39);
            this.txtPassword.TabIndex = 0;
            this.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtPassword.UseSystemPasswordChar = true;
            this.txtPassword.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPassword_KeyDown);
            // 
            // lblTimeLeft
            // 
            this.lblTimeLeft.BackColor = System.Drawing.Color.Black;
            this.lblTimeLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblTimeLeft.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTimeLeft.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.lblTimeLeft.ForeColor = System.Drawing.Color.Chocolate;
            this.lblTimeLeft.Location = new System.Drawing.Point(0, 0);
            this.lblTimeLeft.Name = "lblTimeLeft";
            this.lblTimeLeft.Size = new System.Drawing.Size(545, 31);
            this.lblTimeLeft.TabIndex = 11;
            this.lblTimeLeft.Text = "label3";
            this.lblTimeLeft.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(545, 518);
            this.ControlBox = false;
            this.Controls.Add(this.pnlLogin);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.btRun);
            this.Controls.Add(this.btStop);
            this.Controls.Add(this.cbExcel);
            this.Controls.Add(this.cbWord);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.numericUpDown1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblTimeLeft);
            this.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.RightToLeftLayout = true;
            this.Text = "נעילה אוטומטית - ממשק ניהול";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Resize += new System.EventHandler(this.Form1_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.pnlLogin.ResumeLayout(false);
            this.pnlLogin.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.CheckBox cbWord;
        private System.Windows.Forms.CheckBox cbExcel;
        private System.Windows.Forms.Button btStop;
        private System.Windows.Forms.Button btRun;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel pnlLogin;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Label lblTimeLeft;
    }
}

