/*  Copyright (c) 2016 xanthalas.co.uk
 * 
 *  Author: Xanthalas
 *  Date  : August 2016
 * 
 *  This file is part of ocli
 *
 *  ocli is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  ocli is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with ocli.  If not, see <http://www.gnu.org/licenses/>.
 */

using System;
using System.Collections.Generic;
using System.Linq;
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
        private const string ALIASFILE = "aliases.txt";

        private static Options options = new Options();
        private static List<MailIdentifier> listedEmails = new List<MailIdentifier>();
        private static int mailId = 0;
        private static Aliases aliases;

        static void Main(string[] args)
        {
            if (!CommandLine.Parser.Default.ParseArguments(args, options))
            {
                Console.WriteLine("Invalid arguments passed");
                Console.ReadLine();
                Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
            }

            if (options.Help)
            {
                writeHelp();
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
            string path = AppDomain.CurrentDomain.BaseDirectory;
            aliases = new Aliases(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ALIASFILE));
            var inboxFolder = getInbox();

            if (options.Today)
            {
                options.ShowAll = true;
            }

            var unreadItems = from i in inboxFolder.Items where (i is Outlook.MailItem) && ((Outlook.MailItem)i).UnRead orderby ((Outlook.MailItem)i).ReceivedTime descending select i;
            var allItems = from i in inboxFolder.Items where (i is Outlook.MailItem) orderby ((Outlook.MailItem)i).ReceivedTime descending select i;

            var selectedQuery = unreadItems;

            if (options.ShowAll)
            {
                selectedQuery = allItems;
            }

            int index = 0;

            List<LineData> lines = new List<ocli.LineData>();

            foreach (var item in selectedQuery)
            {
                // not every item in the inbox is a mail item
                Outlook.MailItem mailItem = item as Outlook.MailItem;
                index++;

                var senderName = mailItem.SenderName.Replace(" (CCS)", "").Replace(" (BEU)", "").Trim();
                senderName = (senderName.Length > 20 ? senderName.Substring(0, 20) : senderName);
                senderName = (aliases.NameAlias.ContainsKey(senderName) ? aliases.NameAlias[senderName] : senderName);

                var today = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day);
                var emailDate = new DateTime(mailItem.ReceivedTime.Year, mailItem.ReceivedTime.Month, mailItem.ReceivedTime.Day);

                var age = (today - emailDate).Days;
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

                if (!options.Today || (options.Today && ageString== "[tdy]"))
                {
                    lines.Add(new LineData(index.ToString().PadLeft(3, ' '), ageString, senderName, mailItem.Subject));
                    listedEmails.Add(new ocli.MailIdentifier(index, mailItem.ConversationID, mailItem.ConversationIndex));
                }
            }

            foreach (var line in lines)
            {
                Console.WriteLine($"{line.Id} {line.Age} {line.From.PadRight(LineData.LongestFrom, ' ')} -> {line.Title}");
            }

            saveListedEmailData();
        }

        private static void findAndDisplayEmail(int mailId)
        {
            //First load the email data array
            string line;

            using (StreamReader reader = new StreamReader(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DATAFILE)))
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

            var inboxFolder = getInbox();

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
            string jsonData = JsonConvert.SerializeObject(listedEmails);

            using (StreamWriter writer = new StreamWriter(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DATAFILE)))
            {
                writer.WriteLine(jsonData);
            }

        }

        private static Outlook.MAPIFolder getInbox()
        {
            // start outlook
            Outlook.Application outlookApplication = new Outlook.Application();

            // get inbox 
            Outlook._NameSpace outlookNS = outlookApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            return inboxFolder;
        }

        private static void writeHelp()
        {
            HelpText ht = new HelpText("ocli: Outlook Inbox tool v0.1 (c) Xanthalas 2016");
            ht.AddOptions(options);
            Console.WriteLine(ht.ToString());
            Console.WriteLine("With no parameters it lists unread emails in your Inbox");
            Console.WriteLine("With a single numeric parameter it will open the email from the previous list with the id number given.");
            Console.WriteLine("\nThe list format is: id number [age in days] sender -> title");
        }
    }
}
