namespace TaoFileDoc
{
    partial class Main
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.bntStart = new System.Windows.Forms.Button();
            this.txtWord = new System.Windows.Forms.TextBox();
            this.txtExcel = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(63, 34);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Input Word";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(63, 81);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Input Excel";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // bntStart
            // 
            this.bntStart.Location = new System.Drawing.Point(230, 129);
            this.bntStart.Name = "bntStart";
            this.bntStart.Size = new System.Drawing.Size(75, 23);
            this.bntStart.TabIndex = 2;
            this.bntStart.Text = "Start";
            this.bntStart.UseVisualStyleBackColor = true;
            this.bntStart.Click += new System.EventHandler(this.BntStart_Click);
            // 
            // txtWord
            // 
            this.txtWord.Location = new System.Drawing.Point(159, 37);
            this.txtWord.Name = "txtWord";
            this.txtWord.Size = new System.Drawing.Size(425, 20);
            this.txtWord.TabIndex = 3;
            // 
            // txtExcel
            // 
            this.txtExcel.Location = new System.Drawing.Point(159, 81);
            this.txtExcel.Name = "txtExcel";
            this.txtExcel.Size = new System.Drawing.Size(425, 20);
            this.txtExcel.TabIndex = 4;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.txtExcel);
            this.Controls.Add(this.txtWord);
            this.Controls.Add(this.bntStart);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form Merge Excel to Word";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button bntStart;
        private System.Windows.Forms.TextBox txtWord;
        private System.Windows.Forms.TextBox txtExcel;
    }
}

