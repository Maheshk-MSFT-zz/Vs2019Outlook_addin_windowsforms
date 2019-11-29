using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace templates
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        func f;
        public Ribbon1()
        {
            
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            if(ribbonID=="Microsoft.Outlook.Mail.Compose")
            {
                return GetResourceText("templates.Ribbon2.xml");
            }
            return GetResourceText("templates.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            f = new func();
            
        }
        
        //Callback to create a template. Displays a form which gets the input and in turn calls a method in the func class
        public void create_template(Office.IRibbonControl ctrl)
        {

            Form1 obj = new Form1();
            obj.Show();
            
        }
        
        //Callback to open a template. Calls a method in func class which opens the chosen template.
        public void open(Office.IRibbonControl ctrl)
        {
            
            f.opentemp(ctrl.Id);
        }

        //Callback to edit a template. Calls a method in func class which opens the chosen template.
        public void edit(Office.IRibbonControl ctrl)
        {
            f.edittemp(ctrl.Id);
        }

        //Callback to delete a template. Calls a method in func class which delets the chosen template.
        public void delete(Office.IRibbonControl ctrl)
        {
            f.deletetemp(ctrl.Id);
        }

        //callback to save a template with changes. Calls a method in func class which saves the template.
        public void savetemp(Office.IRibbonControl ctrl)
        {
            f.savetempfunc();
        }

        //Callback to get the contents for the dynamic menu to open a template. 
        //It opens the file containing the template names and creates the dynamic menu with each template as a separate menu.
        public String open_template(Office.IRibbonControl ctrl)
        {
            
            string str;
            string strPath = System.AppDomain.CurrentDomain.BaseDirectory;
            StreamReader rdr = new StreamReader(strPath+"/templates/table.txt");
            string str1="";
            string id1 = "";
            string lbl1 = "";
            int n;
            str = "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">";
            while((str1=rdr.ReadLine())!=null)
            {
                id1 = str1.Substring(0, str1.IndexOf(','));
                n = str1.Length - (str1.IndexOf(',') + 1);
                lbl1 = str1.Substring(str1.IndexOf(',') + 1, n);
                str += "<button id=\"" + id1 + "\" label=\"" + lbl1 + "\" onAction=\"open\" imageMso=\"EnvelopesAndLabelsDialog\"/>";
            }
            
            rdr.Close();
            str += "</menu>";
            return str;
        }

        //Callback to get the contents for the dynamic menu to edit a template. 
        //It opens the file containing the template names and creates the dynamic menu with each template as a separate menu.
        public string edit_template(Office.IRibbonControl ctrl)
        {
            string str;
            string strPath = System.AppDomain.CurrentDomain.BaseDirectory;
            StreamReader rdr = new StreamReader(strPath+"/templates/table.txt");
            string str1 = "";
            string id1 = "";
            string lbl1 = "";
            int n;
            str = "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">";
            while ((str1 = rdr.ReadLine()) != null)
            {
                id1 = str1.Substring(0, str1.IndexOf(','));
                n = str1.Length - (str1.IndexOf(',') + 1);
                lbl1 = str1.Substring(str1.IndexOf(',') + 1, n);
                str += "<button id=\"" + id1 + "\" label=\"" + lbl1 + "\" onAction=\"edit\" imageMso=\"EnvelopesAndLabelsDialog\"/>";
            }
            
            rdr.Close();
            str += "</menu>";
            return str;
        }

        //Callback to delete an existing template.
        //It displays the list of templates and upon choosing one, the chosen one is deleted.
        public string delete_template(Office.IRibbonControl ctrl)
        {
            string str;
            string strPath = System.AppDomain.CurrentDomain.BaseDirectory;
            StreamReader rdr = new StreamReader(strPath + "/templates/table.txt");
            string str1 = "";
            string id1 = "";
            string lbl1 = "";
            int n;
            str = "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">";
            while ((str1 = rdr.ReadLine()) != null)
            {
                id1 = str1.Substring(0, str1.IndexOf(','));
                n = str1.Length - (str1.IndexOf(',') + 1);
                lbl1 = str1.Substring(str1.IndexOf(',') + 1, n);
                str += "<button id=\"" + id1 + "\" label=\"" + lbl1 + "\" onAction=\"delete\" imageMso=\"EnvelopesAndLabelsDialog\"/>";
            }

            rdr.Close();
            str += "</menu>";
            return str;
        }

        //Callback for adding an autocorrect entry. 
        //Displays a form which gets the input and in turn calls a method in func class to add the newly typed entry.
        public void addabbrcallback(Office.IRibbonControl ctrl)
        {
            Form2 obj = new Form2();
            obj.ShowDialog();
        }
        public void removeabbrcallback(Office.IRibbonControl ctrl)
        {
            Form3 obj = new Form3();
            obj.ShowDialog();
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
