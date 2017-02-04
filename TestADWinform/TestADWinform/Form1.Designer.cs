namespace TestADWinform
{
    partial class Form1
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
            this.btn1 = new System.Windows.Forms.Button();
            this.txtBoxname = new System.Windows.Forms.TextBox();
            this.btnOU = new System.Windows.Forms.Button();
            this.delADUser = new System.Windows.Forms.Button();
            this.btnDeleteGroupUser = new System.Windows.Forms.Button();
            this.txtGroupName = new System.Windows.Forms.TextBox();
            this.btnConnect = new System.Windows.Forms.Button();
            this.btnTest = new System.Windows.Forms.Button();
            this.btnMail = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn1
            // 
            this.btn1.Location = new System.Drawing.Point(61, 98);
            this.btn1.Name = "btn1";
            this.btn1.Size = new System.Drawing.Size(75, 23);
            this.btn1.TabIndex = 0;
            this.btn1.Text = "EnterDelAccout";
            this.btn1.UseVisualStyleBackColor = true;
            this.btn1.Click += new System.EventHandler(this.btn1_Click);
            // 
            // txtBoxname
            // 
            this.txtBoxname.Location = new System.Drawing.Point(61, 53);
            this.txtBoxname.Name = "txtBoxname";
            this.txtBoxname.Size = new System.Drawing.Size(100, 21);
            this.txtBoxname.TabIndex = 1;
            // 
            // btnOU
            // 
            this.btnOU.Location = new System.Drawing.Point(182, 98);
            this.btnOU.Name = "btnOU";
            this.btnOU.Size = new System.Drawing.Size(75, 23);
            this.btnOU.TabIndex = 2;
            this.btnOU.Text = "GetOUEntry";
            this.btnOU.UseVisualStyleBackColor = true;
            this.btnOU.Click += new System.EventHandler(this.btnOU_Click);
            // 
            // delADUser
            // 
            this.delADUser.Location = new System.Drawing.Point(304, 97);
            this.delADUser.Name = "delADUser";
            this.delADUser.Size = new System.Drawing.Size(75, 23);
            this.delADUser.TabIndex = 3;
            this.delADUser.Text = "deleteADAccount";
            this.delADUser.UseVisualStyleBackColor = true;
            this.delADUser.Click += new System.EventHandler(this.delADUser_Click);
            // 
            // btnDeleteGroupUser
            // 
            this.btnDeleteGroupUser.Location = new System.Drawing.Point(61, 154);
            this.btnDeleteGroupUser.Name = "btnDeleteGroupUser";
            this.btnDeleteGroupUser.Size = new System.Drawing.Size(100, 23);
            this.btnDeleteGroupUser.TabIndex = 4;
            this.btnDeleteGroupUser.Text = "delUserGroup";
            this.btnDeleteGroupUser.UseVisualStyleBackColor = true;
            this.btnDeleteGroupUser.Click += new System.EventHandler(this.btnDeleteGroupUser_Click);
            // 
            // txtGroupName
            // 
            this.txtGroupName.Location = new System.Drawing.Point(210, 52);
            this.txtGroupName.Name = "txtGroupName";
            this.txtGroupName.Size = new System.Drawing.Size(100, 21);
            this.txtGroupName.TabIndex = 5;
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(181, 153);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(75, 23);
            this.btnConnect.TabIndex = 6;
            this.btnConnect.Text = "connectTion";
            this.btnConnect.UseVisualStyleBackColor = true;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(304, 153);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(75, 23);
            this.btnTest.TabIndex = 7;
            this.btnTest.Text = "测试";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // btnMail
            // 
            this.btnMail.Location = new System.Drawing.Point(73, 207);
            this.btnMail.Name = "btnMail";
            this.btnMail.Size = new System.Drawing.Size(75, 23);
            this.btnMail.TabIndex = 8;
            this.btnMail.Text = "SendMail";
            this.btnMail.UseVisualStyleBackColor = true;
            this.btnMail.Click += new System.EventHandler(this.btnMail_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(433, 262);
            this.Controls.Add(this.btnMail);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.btnConnect);
            this.Controls.Add(this.txtGroupName);
            this.Controls.Add(this.btnDeleteGroupUser);
            this.Controls.Add(this.delADUser);
            this.Controls.Add(this.btnOU);
            this.Controls.Add(this.txtBoxname);
            this.Controls.Add(this.btn1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn1;
        private System.Windows.Forms.TextBox txtBoxname;
        private System.Windows.Forms.Button btnOU;
        private System.Windows.Forms.Button delADUser;
        private System.Windows.Forms.Button btnDeleteGroupUser;
        private System.Windows.Forms.TextBox txtGroupName;
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.Button btnMail;
    }
}

