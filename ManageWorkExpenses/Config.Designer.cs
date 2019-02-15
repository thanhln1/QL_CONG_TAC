namespace ManageWorkExpenses
{
    partial class Config
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose( bool disposing )
        {
            if (disposing && ( components != null ))
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Config));
            this.lblSource = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.tbSource = new System.Windows.Forms.TextBox();
            this.tbUser = new System.Windows.Forms.TextBox();
            this.lblUser = new System.Windows.Forms.Label();
            this.tbPass = new System.Windows.Forms.TextBox();
            this.lblPass = new System.Windows.Forms.Label();
            this.tbDataBase = new System.Windows.Forms.TextBox();
            this.lblDataBase = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(13, 51);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(41, 13);
            this.lblSource.TabIndex = 0;
            this.lblSource.Text = "Source";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(93, 155);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // tbSource
            // 
            this.tbSource.Location = new System.Drawing.Point(93, 51);
            this.tbSource.Name = "tbSource";
            this.tbSource.Size = new System.Drawing.Size(233, 20);
            this.tbSource.TabIndex = 2;
            // 
            // tbUser
            // 
            this.tbUser.Location = new System.Drawing.Point(93, 103);
            this.tbUser.Name = "tbUser";
            this.tbUser.Size = new System.Drawing.Size(233, 20);
            this.tbUser.TabIndex = 4;
            // 
            // lblUser
            // 
            this.lblUser.AutoSize = true;
            this.lblUser.Location = new System.Drawing.Point(13, 103);
            this.lblUser.Name = "lblUser";
            this.lblUser.Size = new System.Drawing.Size(29, 13);
            this.lblUser.TabIndex = 3;
            this.lblUser.Text = "User";
            // 
            // tbPass
            // 
            this.tbPass.Location = new System.Drawing.Point(93, 129);
            this.tbPass.Name = "tbPass";
            this.tbPass.Size = new System.Drawing.Size(233, 20);
            this.tbPass.TabIndex = 6;
            // 
            // lblPass
            // 
            this.lblPass.AutoSize = true;
            this.lblPass.Location = new System.Drawing.Point(13, 129);
            this.lblPass.Name = "lblPass";
            this.lblPass.Size = new System.Drawing.Size(30, 13);
            this.lblPass.TabIndex = 5;
            this.lblPass.Text = "Pass";
            // 
            // tbDataBase
            // 
            this.tbDataBase.Location = new System.Drawing.Point(93, 77);
            this.tbDataBase.Name = "tbDataBase";
            this.tbDataBase.Size = new System.Drawing.Size(233, 20);
            this.tbDataBase.TabIndex = 8;
            // 
            // lblDataBase
            // 
            this.lblDataBase.AutoSize = true;
            this.lblDataBase.Location = new System.Drawing.Point(13, 77);
            this.lblDataBase.Name = "lblDataBase";
            this.lblDataBase.Size = new System.Drawing.Size(54, 13);
            this.lblDataBase.TabIndex = 7;
            this.lblDataBase.Text = "DataBase";
            // 
            // Config
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(344, 195);
            this.Controls.Add(this.tbDataBase);
            this.Controls.Add(this.lblDataBase);
            this.Controls.Add(this.tbPass);
            this.Controls.Add(this.lblPass);
            this.Controls.Add(this.tbUser);
            this.Controls.Add(this.lblUser);
            this.Controls.Add(this.tbSource);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblSource);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Config";
            this.Text = "Config";
            this.Load += new System.EventHandler(this.Config_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox tbSource;
        private System.Windows.Forms.TextBox tbUser;
        private System.Windows.Forms.Label lblUser;
        private System.Windows.Forms.TextBox tbPass;
        private System.Windows.Forms.Label lblPass;
        private System.Windows.Forms.TextBox tbDataBase;
        private System.Windows.Forms.Label lblDataBase;
    }
}