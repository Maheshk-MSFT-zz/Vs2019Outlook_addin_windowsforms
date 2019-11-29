using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using outlook = Microsoft.Office.Interop.Outlook;

namespace templates
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //To create a new template.
        private void Create_Click(object sender, EventArgs e)
        {
            string fn;
            if(textBox1.Text.Equals(""))
            {
                MessageBox.Show("Please enter the File name","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else
            {
                fn=textBox1.Text;
                int flag = 0;
                string strPath = System.AppDomain.CurrentDomain.BaseDirectory;
                StreamReader rdr = new StreamReader(strPath+"/templates/table.txt");
                string str, substr;
                while(rdr.EndOfStream!=true)
                {
                    str = rdr.ReadLine();
                    substr = str.Substring(str.IndexOf(',')+1);
                    if(fn.Equals(substr))
                    {
                        flag = 1;
                        break;
                    }
                }
                rdr.Close();
                //The typed in name is checked with the existing template names.
                if (flag == 1)
                {
                    MessageBox.Show("A template already exists with the mentioned name. Change template name completely!!","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
                else
                {
                    
                    MessageBox.Show("Type the contents of your template and click save template in the Templates tab to save the newly created template","Information",MessageBoxButtons.OKCancel,MessageBoxIcon.Information);
                    this.Close();
                    func f = new func();
                    f.createmi(fn);
                    
                }
                
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
