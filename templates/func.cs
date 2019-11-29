using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using word = Microsoft.Office.Interop.Word;
using System.IO;

namespace templates
{
    class func
    {
        Outlook.Application oapp;
        public static string fn="";
        string strPath = System.AppDomain.CurrentDomain.BaseDirectory;
        public func()
        {
            oapp = Globals.ThisAddIn.Application;
        }
        //To create a new template. The function is called withe the typed in template name as a parameter.
        //The files are created with the file name "id*" where * denotes a sequence number. For eg, id1 denotes the first template.
        //The file is stored in .msg format.
        public void createmi(String fn1)
        {
            string res,ln="",substr,id;
            string[] split1;
            int num;
            StreamReader rdr = new StreamReader(strPath+"/templates/table.txt");
            while (rdr.EndOfStream !=true)
            { 
                ln = rdr.ReadLine();
            }
            rdr.Close();
            if(ln.Length==0)
                id="id1";
            else
            {
                split1 = ln.Split(',');
                substr = split1[0].Substring(2);
                num=Convert.ToInt32(substr);
                num = num + 1;
                id = "id" + num.ToString();
            }
            Outlook.MailItem omail;
            omail = oapp.CreateItem(Outlook.OlItemType.olMailItem);
            omail.SaveAs(strPath+"/templates/" + id + ".msg");
            StreamWriter wdr = new StreamWriter(strPath+"/templates/table.txt", true);
            res = id + "," + fn1;
            wdr.WriteLine(res);
            wdr.Close();
            fn = id;
            omail.Display(true);  
        }

        //To open a template.
        public void opentemp(string id)
        {
            fn = id;
            Outlook.MailItem omail;
            omail = oapp.CreateItemFromTemplate(strPath+"/templates/" + id + ".msg");
            omail.Display(true);
        }

        //To save the changes made in the template after editing.
        public void savetempfunc()
        {
            Outlook.MailItem omail;
            omail=oapp.ActiveInspector().CurrentItem;
            if (fn.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("Cannot save template without file name. Please add templates using 'CREATE TEMPLATE' option","Error",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
                omail.Close(Outlook.OlInspectorClose.olDiscard);
            }
            else
            {
                omail.SaveAs(strPath + "/templates/" + fn + ".msg", System.Reflection.Missing.Value);
                omail.Close(Outlook.OlInspectorClose.olDiscard);
                omail = oapp.CreateItemFromTemplate(strPath + "/templates/" + fn + ".msg");
                omail.HTMLBody = RemoveSignature(omail.HTMLBody);
                omail.SaveAs(strPath + "/templates/" + fn + ".msg", System.Reflection.Missing.Value);
                fn = "";
            }
        }

        //To open the template for editing.
        public void edittemp(string id)
        {
            fn=id;
            Outlook.MailItem omail;
            omail = oapp.CreateItemFromTemplate(strPath+"/templates/" + id + ".msg");
            omail.Display(true);
        }

        //To delete a template
        public void deletetemp(string id)
        {
            System.Windows.Forms.DialogResult dr;
            dr = System.Windows.Forms.MessageBox.Show("Do you want to delete the template?", "Confirmation", System.Windows.Forms.MessageBoxButtons.OKCancel,System.Windows.Forms.MessageBoxIcon.Question); 
            if (dr.ToString().Equals("OK"))
            {
                StreamReader rdr = new StreamReader(strPath + "/templates/table.txt");
                StringBuilder bdr = new StringBuilder();
                string str, id1 = "";
                int pos1, pos2;
                while (!(rdr.EndOfStream))
                {
                    str = rdr.ReadLine();
                    pos1 = 0;
                    pos2 = str.IndexOf(',');
                    if (str.Substring(pos1, pos2).Equals(id))
                    {
                        id1 = id;
                        continue;
                    }
                    bdr.AppendLine(str);
                }
                rdr.Close();
                StreamWriter wdr = new StreamWriter(strPath + "/templates/table.txt");
                wdr.Write(bdr);
                wdr.Close();
                if (File.Exists(strPath + "/templates/" + id1 + ".msg"))
                {
                    File.Delete(strPath + "/templates/" + id1 + ".msg");
                }
            }
        } 
        //To add the autocorrect entry into the autocorrect list.
        public void handleabbr(string abbr,string expansion)
        {
            Outlook.MailItem omail = oapp.CreateItem(Outlook.OlItemType.olMailItem);
            Outlook.Inspector ins;
            ins = omail.GetInspector;
            word.Document doc;
            word.Application wapp;
            doc = ins.WordEditor;
            wapp = doc.Application;
            word.AutoCorrectEntries ace;
            ace = wapp.AutoCorrect.Entries;
            ace.Add(abbr, expansion);
            System.Windows.Forms.MessageBox.Show("New entry added","Information",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Information);

        }

        //To delete an auto correct entry
        public void deleteabbr(string abbr)
        {
            
            System.Windows.Forms.DialogResult dr;
            int flag = 0;
            Outlook.MailItem omail = oapp.CreateItem(Outlook.OlItemType.olMailItem);
            Outlook.Inspector ins;
            ins = omail.GetInspector;
            word.Document doc;
            word.Application wapp;
            doc = ins.WordEditor;
            wapp = doc.Application;
            word.AutoCorrectEntries ace;
            ace = wapp.AutoCorrect.Entries;
            foreach (word.AutoCorrectEntry ace1 in ace)
            {
                if (ace1.Name.Equals(abbr))
                {
                    dr = System.Windows.Forms.MessageBox.Show("Do you want to delete the entry?", "Confirmation", System.Windows.Forms.MessageBoxButtons.OKCancel,System.Windows.Forms.MessageBoxIcon.Question);
                    flag = 1;
                    if (dr.ToString().Equals("OK"))
                    {
                        ace1.Delete();
                    }  
                }
            }
            if(flag==0)
            {
                System.Windows.Forms.MessageBox.Show("Entry not found","Error",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
            }
            
        }
        
        //To write the default template names into the "table" text file.
        public void writedefault(string[] fn1)
        {
            Outlook.MailItem omail;
            omail = oapp.CreateItem(Outlook.OlItemType.olMailItem);
            
            string id="",res;
            StreamWriter wdr = new StreamWriter(strPath + "/templates/table.txt", true); 
            for (int i = 0; i <2 ; i++)
            {
                id = "id" + i.ToString();
                omail.SaveAs(strPath + "/templates/" + id + ".msg");
                res = id + "," + fn1[i];
                wdr.WriteLine(res);
            }
            wdr.Close();
        }

        //To create the default templates.
        public void writecont(string fn1)
        {
            Outlook.MailItem omail; 
                if(fn1.Equals("Opening template"))
                {
                    omail = oapp.CreateItemFromTemplate(strPath + "/templates/id0.msg");
                    omail.Body = "Hello <customer_name>>,\n\nMy name is <<Engineer name>> and I work with Visual Studio Office Development team. I have taken the ownership of your case # SRQ080602600318 where you are facing an issue in publishing VSTO (Visual Studio Tools for Office) based PowerPoint add-in.\n\nAccording to the case notes,\n\nI would like to discuss this issue further with you.\n\nPlease let me know the convenient time for you, to discuss the issue further.";
                    omail.SaveAs(strPath + "/templates/id0.msg");
                }
                if (fn1.Equals("Closing template"))
                {
                    omail = oapp.CreateItemFromTemplate(strPath + "/templates/id1.msg");
                    omail.Body = "Hello\n\nIt was my pleasure to assist you with your issue. I am providing you with a summary of the key points of the case for your records. If you have any questions please feel free to contact me; my contact information is listed below.\n\nBased on our last communication, I will close the case.Also please let my manager, Saumya Dasgupta(saumya.dasgupta@microsoft.com) know how you feel about the support experience provided to you for this incident.\n\nProblem - \nEnvironment - \nTroubleshooting -\nRoot Cause (if known) - \nResolution - \nRelated Knowledge Base Articles\n=======================\n\nAdditional Information and Recommendations\n=================================\n\n\nThank you for choosing Microsoft.";
                    omail.SaveAs(strPath + "/templates/id1.msg");
                    
                }
              
            
        }
        
        //To remove signature befor saving template.
        public static string RemoveSignature(string strHTMLBODY)
        {

            string strStartPTag = "<p class=";
            string strSearTag = "style='mso-bookmark:_MailAutoSig'";
            string strEndPTag = "</p>";
            try
            {
                int TagIndex = strHTMLBODY.IndexOf(strSearTag, 0);
                int startPTagIndex = 0, endPTagIndex = 0, PosNavigator = 0;
                while (TagIndex > 0)
                {
                    startPTagIndex = 0; endPTagIndex = 0; PosNavigator = 0;

                    PosNavigator = TagIndex;

                    while (PosNavigator > 0)
                    {
                        if (strHTMLBODY.Substring(PosNavigator, strStartPTag.Length) == strStartPTag)
                        {
                            startPTagIndex = PosNavigator;
                            break;
                        }
                        PosNavigator--;
                    }

                    endPTagIndex = strHTMLBODY.IndexOf(strEndPTag, TagIndex);

                    if (startPTagIndex > 0 && endPTagIndex > 0)
                    {
                        strHTMLBODY = strHTMLBODY.Remove(startPTagIndex, endPTagIndex - startPTagIndex + strEndPTag.Length);
                    }

                    TagIndex = strHTMLBODY.IndexOf(strSearTag, 0);

                }

                
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            return strHTMLBODY;
        }

        
    }
}