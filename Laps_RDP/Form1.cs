using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laps_RDP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string Output = getDomaininfo();
                domainName = Output;
            }
            catch (Exception)
            {

                throw;
            }
        }

        string pc = string.Empty;
        string username = @"administrator";
        string domainName ;
        string pcDomain = string.Empty;

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            pc = textBox1.Text;
            pcDomain = pc + "." + domainName;
            string user = pcDomain + "\\" + username;

            if (pc != null)
            {
                try
                {
                    DomainController domainController = DomainController.FindOne(new DirectoryContext(DirectoryContextType.Domain));
                    string fqdnServerName = domainController.Name;
                    string[] dcsplit = domainController.Name.Split('.');
                    string dcnetbios = dcsplit[0];
                    string[] domainComponents = domainName.Split('.');
                    string ldapPath = $"LDAP://{dcnetbios}/{string.Join(",", domainComponents.Select(dc => "dc=" + dc))}";
                    
                    using (DirectoryEntry entry = new DirectoryEntry(ldapPath))
                    {
                        using (DirectorySearcher adSearcher = new DirectorySearcher(entry))
                        {
                            string computerName = pc;
                            adSearcher.Filter = "(&(objectClass=computer)(cn=" + computerName + "))";
                            adSearcher.SearchScope = SearchScope.Subtree;
                            adSearcher.PropertiesToLoad.Add("description");
                            SearchResult searchResult = adSearcher.FindOne();

                            try
                            {
                                if (searchResult != null)
                                {
                                    string ip = string.Empty;
                                    PropertyValueCollection pass = searchResult.GetDirectoryEntry().Properties["ms-Mcs-AdmPwd"];
                                    string result = (pass != null && pass.Value != null) ? pass.Value.ToString() : "";
                                    if (result != string.Empty)
                                    {
                                        var process = new Process();
                                        {
                                            process.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
                                            process.StartInfo.Arguments = $"/generic:\"{pcDomain}\" /user:\"{user}\" /pass:\"{result}\"";
                                            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                                            process.Start();
                                        }

                                        // initiate RDP with Saved Creds
                                        process.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
                                        process.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
                                        process.StartInfo.Arguments = String.Format("/v:{0}", pcDomain);
                                        process.Start();

                                        // Tidy up saved credentials
                                        //process.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
                                        //process.StartInfo.Arguments = $"/delete:\"{pcDomain}\"";
                                        //process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                                        //process.Start();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Computer found " + pc + "but no LAPS password in record");
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Unable to find computer");
                                }

                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message.ToString());
                }

            }
            else
            {
                MessageBox.Show("Please enter a PCID");
            }
            button1.Enabled = true;
        }

        public string getDomaininfo()
        {
            Domain domain = Domain.GetCurrentDomain();
            string domainName = domain.Name;
            return domainName;
        }
    }
}
