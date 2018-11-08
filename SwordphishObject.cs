using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Text;
using System;
using System.Security.Authentication;

namespace NotifySecurity
{
    static class SwordphishObject
    {

        static string SwordphishURL = Properties.Settings.Default.SwordphishURL;
        static string URLrequest = SwordphishURL + "/result/report/ID";
        static string SworphishHeader = Properties.Settings.Default.SwordPhishHeader;
        static string WebExpID = SworphishHeader + @": \[[0-9a-z-]+\]";
        static string WebExpPrefix = SworphishHeader + @": [";
        static string WebExpSuffix = @"]";

        public static string NoHeaderFound = "no header found";
        public static string AnswerFromSwordphish = "Answer from wordphish server: ";
        public static string NoAnswerFromSwordphish = "NO ANSWER from Swordphish server";

        public static string MsgIfSwordphishDetected= "\n\nWell done, you've well identified our fake mail generated for the phishing campain !";


        public static string SetHeaderIDtoURL(string headers)
        {
            var pattern = WebExpID;
            var regex = new Regex(pattern);
            var match = regex.Match(headers);
            foreach (var group in match.Groups)
            {
                if(group.ToString().Trim()!=string.Empty)
                { 
                    //we got the ID : fill in the URL
                    string sURL = URLrequest
                        .Replace(@"/ID", string.Concat(@"/" + group.ToString()))
                        .Replace(WebExpPrefix, string.Empty)
                        .Replace(WebExpSuffix, string.Empty);
                    return sURL;
                }
            }

            return NoHeaderFound;

        }

        public const SslProtocols _strTls12 = (SslProtocols)0x00000C00;
        public const SecurityProtocolType Tls12 = (SecurityProtocolType)_strTls12;

        public static string SendNotification(string sURL)
        {

            ServicePointManager.SecurityProtocol = Tls12;
            
            string strToWriteNOK = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + ":" + NoAnswerFromSwordphish + "\n(" + sURL + ")";

            try
            {

               
                string html = string.Empty;

               var request = (HttpWebRequest)WebRequest.Create(sURL);            
                var response = (HttpWebResponse)request.GetResponse();
                html = new StreamReader(response.GetResponseStream()).ReadToEnd();


                IWebProxy defaultProxy = WebRequest.DefaultWebProxy;
                if (defaultProxy != null)
                {
                    defaultProxy.Credentials = CredentialCache.DefaultCredentials;                }
                WebClient client = new WebClient();
                client.Proxy = defaultProxy;
                html = client.DownloadString(sURL);



                string strAnswerFromSwordphish = html;
                string strToWriteOK = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + ":" + AnswerFromSwordphish + strAnswerFromSwordphish;

               
                return strAnswerFromSwordphish;
            }
            catch (System.Exception exc)
            {
                strToWriteNOK = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + ":" + "\n(" + sURL + ")\n" + exc.ToString();
            }

            return NoAnswerFromSwordphish;
        }
    }
}

