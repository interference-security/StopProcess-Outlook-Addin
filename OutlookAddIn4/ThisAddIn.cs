using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Management;
using System.ComponentModel;
using System.Configuration.Install;
using System.Collections;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Security.Cryptography;

namespace OutlookAddIn4
{
    [RunInstaller(true)]
    public class InstallHelper : Installer
    {
        // Override the 'Install' method of the Installer class.
        public override void Install( IDictionary mySavedState )
        {
            base.Install( mySavedState );
            // Code maybe written for installation of an application.
            string uniq_val = "saU5zd60)Sm8ghIGnD{V7{[eu(vR0w{vXBuHl9Q}<VO2y{YLTM5mk8Bk<0oVUX4S";
            RegistryKey r = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\Outlook\Addins\StopSendProcess");
            string reg_key_val = ((Context.Parameters["ProcessName"]).ToString()).ToLower();
            reg_key_val = sha256_hash(uniq_val + reg_key_val).ToLower();
            r.SetValue("MyProcess", reg_key_val);
        }

        public static String sha256_hash(String value)
        {
            StringBuilder Sb = new StringBuilder();
            using (SHA256 hash = SHA256Managed.Create())
            {
                Encoding enc = Encoding.UTF8;
                Byte[] result = hash.ComputeHash(enc.GetBytes(value));
                foreach (Byte b in result)
                    Sb.Append(b.ToString("x2"));
            }
            return Sb.ToString();
        }
    }

    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {
            //System.Windows.Forms.MessageBox.Show("Your Ribbon Works!");
            string uniq_val = "saU5zd60)Sm8ghIGnD{V7{[eu(vR0w{vXBuHl9Q}<VO2y{YLTM5mk8Bk<0oVUX4S";
            RegistryKey r = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\Outlook\Addins\StopSendProcess");

            String ProcessName = ((r.GetValue("MyProcess")).ToString()).ToLower();
            int StopSend = 1;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Process");

            foreach (ManagementObject queryObj in searcher.Get())
            {
                String wmi_process = (String)queryObj["Name"];
                String wmi_process_hash = sha256_hash(uniq_val + wmi_process.ToLower()).ToLower();
                if (wmi_process_hash == ProcessName)
                {
                    StopSend = 0;
                    break;
                }
            }

            if (StopSend == 0)
            {
                Cancel = false;
                //MessageBox.Show("Success");
            }
            else
            {
                Cancel = true;
                //MessageBox.Show("Fail");
            }
        }

        public static String sha256_hash(String value)
        {
            StringBuilder Sb = new StringBuilder();
            using (SHA256 hash = SHA256Managed.Create())
            {
                Encoding enc = Encoding.UTF8;
                Byte[] result = hash.ComputeHash(enc.GetBytes(value));
                foreach (Byte b in result)
                    Sb.Append(b.ToString("x2"));
            }
            return Sb.ToString();
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
