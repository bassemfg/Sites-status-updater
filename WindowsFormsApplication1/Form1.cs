using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

namespace WindowsFormsApplication1
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            string siteUrl = "https://hp-27b0ee14ded081.sharepoint.com/teams/spohub/ACSMigrationManager/";

            ClientContext clientContext = new ClientContext(siteUrl);
            System.Security.SecureString pwd = new System.Security.SecureString();
            pwd.AppendChar('p');
            pwd.AppendChar('a');
            pwd.AppendChar('s');
            pwd.AppendChar('s');
            pwd.AppendChar('@');
            pwd.AppendChar('w');
            pwd.AppendChar('o');
            pwd.AppendChar('r');
            pwd.AppendChar('d');
            pwd.AppendChar('1');
            clientContext.Credentials = new SharePointOnlineCredentials("bassem.yacoube@hp.com", pwd);
            Web site = clientContext.Web;
            clientContext.Load(site);
            clientContext.ExecuteQuery();

            SP.List oList = clientContext.Web.Lists.GetByTitle("Migration Tasks");
            CamlQuery query;
            string sitesText = "" + textBox1.Text;
            sitesText = sitesText.Replace("\r", "");
            sitesText = sitesText.Replace("\n", ",");
            string[] sites = null;
            if (sitesText.Length > 0)
            {
                sites = sitesText.Split(',');

                for (int i = 0; i < sites.Length; i++)
                {
                    if (sites[i].Trim().Length > 0)
                    {
                        query = new CamlQuery();
                        query.ViewXml = "<View><Query><Where><Contains><FieldRef Name='ContentSource'/><Value Type='Text'>" +
                            sites[i] + "</Value></Contains></Where></Query></View>";
                        ListItemCollection collListItem = oList.GetItems(query);

                        clientContext.Load(collListItem);
                        clientContext.ExecuteQuery();



                        if (collListItem.Count == 1)
                        {
                            ListItem oListItem = collListItem[0];
                            //listBox1.DataSource = collListItem;
                            textBox3.Text += oListItem["Title"].ToString() + @"
";
                            oListItem["MigrationStatus"] = textBox2.Text;
                            oListItem.Update();
                            clientContext.ExecuteQuery();
                        }
                    }
                }
            }
        }
    }
}

