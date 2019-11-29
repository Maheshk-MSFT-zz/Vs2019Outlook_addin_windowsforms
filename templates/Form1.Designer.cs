namespace templates
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Create = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.Cancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(143, 6);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(412, 20);
            this.textBox1.TabIndex = 0;
            // 
            // Create
            // 
            this.Create.Location = new System.Drawing.Point(324, 67);
            this.Create.Name = "Create";
            this.Create.Size = new System.Drawing.Size(112, 30);
            this.Create.TabIndex = 2;
            this.Create.Text = "Create a template";
            this.Create.UseVisualStyleBackColor = true;
            this.Create.Click += new System.EventHandler(this.Create_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(125, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Enter the template name:";
            // 
            // Cancel
            // 
            this.Cancel.Location = new System.Drawing.Point(442, 67);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(113, 30);
            this.Cancel.TabIndex = 5;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(140, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(369, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "For eg: An \"opening mail\" template can have its name as \"opening template\"";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(563, 104);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Create);
            this.Controls.Add(this.textBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "Templates";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button Create;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button Cancel;
        private System.Windows.Forms.Label label1;
    }
}