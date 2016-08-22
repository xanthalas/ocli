using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using CommandLine.Text;
using System.IO;
using Newtonsoft.Json;

namespace ocli
{
    class Program
    {
        private const string DATAFILE = "ocli.dat";

        private static Options options = new Options();
        private static List<MailIdentifier> listedEmails = new List<MailIdentifier>();
        private static int mailId = 0;

        static void Main(string[] args)
        {
            if (!CommandLine.Parser.Default.ParseArguments(args, options))
            {
                Console.WriteLine("Invalid arguments passed");
                Console.ReadLine();
                Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
            }

            if (args.Length == 1 && !args[0].StartsWith("-"))
            {
                bool worked = int.TryParse(args[0], out mailId);
            }

            if (mailId > 0)
            {
                findAndDisplayEmail(mailId);
            }
            else
            {
                listEmails();
            }

            //Console.ReadLine();
        }

        private static void listEmails()
        {
            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();

            // get inbox 
            Outlook._NameSpace outlookNS = outlookApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            var unreadItems = from i in inboxFolder.Items where (i is Outlook.MailItem) && ((Outlook.MailItem)i).UnRead orderby ((Outlook.MailItem)i).ReceivedTime descending select i;
            var allItems = from i in inboxFolder.Items where (i is Outlook.MailItem) orderby ((Outlook.MailItem)i).ReceivedTime descending select i;

            var selectedQuery = unreadItems;

            if (options.ShowAll)
            {
                selectedQuery = allItems;
            }

            int index = 0;

            foreach (var item in selectedQuery)
            {
                // not every item in the inbox is a mail item
                Outlook.MailItem mailItem = item as Outlook.MailItem;
                index++;

                var senderName = mailItem.SenderName.Replace(" (CCS)", "");
                senderName = senderName.Replace(" (BEU)", "");
                senderName = (senderName.Length > 20 ? senderName.Substring(0, 20) : senderName.PadRight(20, ' '));

                var age = (DateTime.Today - mailItem.ReceivedTime).Days;
                string ageString = "";
                if (age == 0)
                {
                    ageString = "[tdy]";
                }
                else
                {
                    ageString = "-" + age.ToString();
                    ageString = "[" + ageString.PadLeft(3, ' ') + "]";
                }

                Console.WriteLine($"{index.ToString().PadLeft(3, ' ')} {ageString} {senderName} -> {mailItem.Subject}");

                listedEmails.Add(new ocli.MailIdentifier(index, mailItem.ConversationID, mailItem.ConversationIndex));
            }

            saveListedEmailData();
        }

        private static void findAndDisplayEmail(int mailId)
        {
            //First load the email data array
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string line;

            using (StreamReader reader = new StreamReader(path + @"\" + DATAFILE))
            {
                line = reader.ReadLine();
            }

            if (line != null && line.Length > 0)
            {
                listedEmails =  Newtonsoft.Json.JsonConvert.DeserializeObject<List<MailIdentifier>>(line);
            }

            if (listedEmails == null || listedEmails.Count == 0)
            {
                return;
            }

            var selectedEmail = (from l in listedEmails where l.Id == mailId select l).FirstOrDefault();

            if (selectedEmail == null)
            {
                Console.WriteLine($"There is no email with id {mailId}");
                return;
            }


            //Now find the selected email
            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();

            // get inbox 
            Outlook._NameSpace outlookNS = outlookApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            var requiredEmail = (from i in inboxFolder.Items where (i is Outlook.MailItem) 
                                && ((Outlook.MailItem)i).ConversationID == selectedEmail.ConversationId
                                && ((Outlook.MailItem)i).ConversationIndex == selectedEmail.ConversationIndex
                                select i).FirstOrDefault();

            if (requiredEmail != null)
            {
                ((Outlook.MailItem)requiredEmail).Display();
            }
        }

        private static void saveListedEmailData()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string jsonData = JsonConvert.SerializeObject(listedEmails);

            using (StreamWriter writer = new StreamWriter(path + @"\" + DATAFILE))
            {
                writer.WriteLine(jsonData);
            }

        }
    }
}
