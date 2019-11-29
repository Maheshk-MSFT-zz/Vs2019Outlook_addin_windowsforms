using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;


namespace templates
{
    public partial class ThisAddIn
    {
        
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Cheks if the folder/file exists or not and if not creates them
            string strPath = System.AppDomain.CurrentDomain.BaseDirectory+"/templates";
            string fname=System.AppDomain.CurrentDomain.BaseDirectory+"/templates/table.txt";
            FileStream fs=null;
            if(!(Directory.Exists(strPath)))
            {
                Directory.CreateDirectory(strPath);
                fs = new FileStream(fname, FileMode.CreateNew);
                fs.Close();

            }
            else if(!(File.Exists(fname)))
            {
                fs = new FileStream(fname, FileMode.CreateNew);
                fs.Close();
            }
            string[] fn = { "Opening template", "Closing template" };
            string temp1 = System.AppDomain.CurrentDomain.BaseDirectory + "/templates/id0.msg";
            string temp2 = System.AppDomain.CurrentDomain.BaseDirectory + "/templates/id1.msg";
            if (!(File.Exists(temp1))&&(!(File.Exists(temp2))))
            {
                func f = new func();
                f.writedefault(fn);
                f.writecont(fn[0]);
                f.writecont(fn[1]);
            }
            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
