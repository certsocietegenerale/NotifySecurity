using NotifySecurity.Properties;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace NotifySecurity
{


    [ComVisible(true)]
    public class Ribbon1 : IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public Ribbon1()
        {
            StartUp = true;
        }

        public Boolean StartUp = false;
        public String ddlEntityValue = "Company";

        #region IRibbonExtensibility Members


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

        public string GetCustomUI(string ribbonID)
        {

            String txtRibbon = GetResourceText("NotifySecurity.Ribbon1.xml");

            return txtRibbon;
        }

        #region Ribbon Callbacks
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {

            this.ribbon = ribbonUI;
        }

        public Bitmap Btn_GetImage(IRibbonControl control)
        {

            return new Bitmap(Resources.shieldy);
        }
        #endregion


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

            try
            {
                Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo("Company");
            }
            catch (System.Exception)
            {
            }

        }


        public string GetContextMenuLabel(Office.IRibbonControl control)
        {
            return "Alert Security";
        }

        public string GetGroupLabel(Office.IRibbonControl control)
        {
            return "Shieldy";

        }
        public string GetTabLabel(Office.IRibbonControl control)
        {
            return "Alert Security";

        }


        public string GetSupertipLabel(Office.IRibbonControl control)
        {
            var v = System.Reflection.Assembly.GetAssembly(typeof(Ribbon1)).GetName().Version;
            int revMaj = v.Major;
            int revMin = v.Minor;
            int revBuild = v.Build;
            int revRev = v.Revision;

            return "Shieldy v" + revMaj.ToString() + "." + revMin.ToString() + "." + revBuild.ToString() + "." + revRev.ToString();// versionInfo.ToString();
        }


        public string GetScreentipLabel(Office.IRibbonControl control)
        {
            return "Alert Security";
        }

        public string GetButtonLabel(Office.IRibbonControl control)
        {
            return "Alert Security";
        }

        public void ShowMessageClick(Office.IRibbonControl control)
        {

            CreateNewMailToSecurityTeam(control);
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.shieldy);

        }

        private void CreateNewMailToSecurityTeam(IRibbonControl control)
        {

            Selection selection =
                Globals.ThisAddIn.Application.ActiveExplorer().Selection;

            if (selection.Count == 1)   // Check that selection is not empty.
            {
                object selectedItem = selection[1];   // Index is one-based.
                Object mailItemObj = selectedItem as Object;
                MailItem mailItem = null;// selectedItem as MailItem;
                if (selection[1] is Outlook.MailItem)
                {
                    mailItem = selectedItem as MailItem;
                }

                MailItem tosend = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                tosend.Attachments.Add(mailItemObj);

                #region create mail from default
                try
                {

                    tosend.To = Properties.Settings.Default.Security_Team_Mail;
                    tosend.Subject = "[User Alert] Suspicious mail";

                    tosend.CC = Properties.Settings.Default.Security_Team_Mail_cc;
                    tosend.BCC = Properties.Settings.Default.Security_Team_Mail_bcc;

                    #region retrieving message header
                    string allHeaders = "";
                    if (selection[1] is Outlook.MailItem)
                    {
                        string[] preparedByArray = mailItem.Headers("X-PreparedBy");
                        string preparedBy;
                        if (preparedByArray.Length == 1)
                            preparedBy = preparedByArray[0];
                        else
                            preparedBy = "";
                        allHeaders = mailItem.HeaderString();
                    }
                    else
                    {
                        string typeFound = "unknown";
                        typeFound = (selection[1] is Outlook.MailItem) ? "MailItem" : typeFound;

                        if (typeFound == "unknown")
                            typeFound = (selection[1] is Outlook.MeetingItem) ? "MeetingItem" : typeFound;

                        if (typeFound == "unknown")
                            typeFound = (selection[1] is Outlook.ContactItem) ? "ContactItem" : typeFound;

                        if (typeFound == "unknown")
                            typeFound = (selection[1] is Outlook.AppointmentItem) ? "AppointmentItem" : typeFound;

                        if (typeFound == "unknown")
                            typeFound = (selection[1] is Outlook.TaskItem) ? "TaskItem" : typeFound;

                        allHeaders = "Selected Outlook item was not a mail (" + typeFound + "), no header extracted";
                    }

                    #endregion

                    string SwordPhishURL = SwordphishObject.SetHeaderIDtoURL(allHeaders);

                    if (SwordPhishURL != SwordphishObject.NoHeaderFound)
                    {
                        string SwordPhishAnswer = SwordphishObject.SendNotification(SwordPhishURL);
                    }
                    else
                    {
                        tosend.Body = "Hello, I received the attached email and I think it is suspicious";
                        tosend.Body += "\n";
                        tosend.Body += "I think this mail is malicious for the following reasons:";
                        tosend.Body += "\n";
                        tosend.Body += "Please analyze and provide some feedback.";
                        tosend.Body += "\n";
                        tosend.Body += "\n";

                        tosend.Body += GetCurrentUserInfos();

                        tosend.Body += "\n\nMessage headers: \n--------------\n" + allHeaders + "\n\n";

                        tosend.Save();
                        tosend.Display();
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Using default template" + ex.Message);

                    MailItem mi = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                    mi.To = Properties.Settings.Default.Security_Team_Mail;
                    mi.Subject = "Security addin error";
                    String txt = ("An error occured, please notify your security contact and give him/her the following information: " + ex);
                    mi.Body = txt;
                    mi.Save();
                    mi.Display();
                }
            }
            else if (selection.Count < 1)   // Check that selection is not empty.
            {
                MessageBox.Show("Please select one mail.");
            }
            else if (selection.Count > 1)
            {
                MessageBox.Show("Please select only one mail to be raised to the security team.");
            }
            else
            {
                MessageBox.Show("Bad luck... this case has not been identified by the dev");
            }


        }
        #endregion


        public String GetCurrentUserInfos()
        {

            String wComputername = System.Environment.MachineName + " (" + System.Environment.OSVersion.ToString() + ")";
            String wUsername = System.Environment.UserDomainName + "\\" + System.Environment.UserName;

            string str = "Possibly useful information:\n--------------";


            Outlook.AddressEntry addrEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                Outlook.ExchangeUser currentUser =
                    Globals.ThisAddIn.Application.Session.CurrentUser.
                    AddressEntry.GetExchangeUser();
                if (currentUser != null)
                {
                    str += "\n - Name: " + currentUser.Name;
                    str += "\n - STMP address: " + currentUser.PrimarySmtpAddress;
                    str += "\n - Title: " + currentUser.JobTitle;
                    str += "\n - Department: " + currentUser.Department;
                    str += "\n - Location: " + currentUser.OfficeLocation;
                    str += "\n  - Business phone: " + currentUser.BusinessTelephoneNumber;
                    str += "\n - Mobile phone: " + currentUser.MobileTelephoneNumber;

                }
            }
            str += "\n - Windows username:" + wUsername;
            str += "\n - Computername:" + wComputername;
            str += "\n";
            return str;
        }

    }

    public static class MailItemExtensions
    {
        private const string HeaderRegex =
            @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)" +
                "(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";
        private const string TransportMessageHeadersSchema =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static string[] Headers(this MailItem mailItem, string name)
        {
            var headers = mailItem.HeaderLookup();
            if (headers.Contains(name))
                return headers[name].ToArray();
            return new string[0];
        }

        public static ILookup<string, string> HeaderLookup(this MailItem mailItem)
        {
            var headerString = mailItem.HeaderString();
            var headerMatches = Regex.Matches
                (headerString, HeaderRegex, RegexOptions.Multiline).Cast<Match>();
            return headerMatches.ToLookup(
                h => h.Groups["header_key"].Value,
                h => h.Groups["header_value"].Value);
        }

        public static string HeaderString(this MailItem mailItem)
        {
            return (string)mailItem.PropertyAccessor
                .GetProperty(TransportMessageHeadersSchema);
        }

    }

}
#endregion