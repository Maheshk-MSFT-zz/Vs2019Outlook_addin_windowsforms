using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace templates
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //To add an autocorrect entry
        private void Add_Click(object sender, EventArgs e)
        {
            if(textBox1.Text.Equals("") || textBox2.Text.Equals(""))
            {
                MessageBox.Show("Please enter Abbreviation/expansion","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else
            {
                string ab, ex;
                ab = textBox1.Text;
                ex = textBox2.Text;
                func f1 = new func();
                f1.handleabbr(ab, ex);
                this.Close();
            }
        }

        

        

    }
}
