using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace templates
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Delete_Click(object sender, EventArgs e)
        {
            if(textBox1.Text.Length==0)
            {
                MessageBox.Show("Please enter the entry to be deleted", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                func f =new func();
                f.deleteabbr(textBox1.Text);
                this.Close();
                
            }
        }
    }
}
