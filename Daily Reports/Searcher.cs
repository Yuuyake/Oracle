using System;
using System.Runtime.InteropServices;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Daily_Reports {
    public class Searcher {
        public string advancedSearchTag = "Our first advanced search in Outlook";
        public Outlook.Search advancedSearch;
        public Outlook.Application OApp = new Outlook.Application();
        public bool isFinished = false;

        public Outlook.Search RunAdvancedSearch(string wordInSubject) {
            string scope = "Inbox";
            string filter = String.Format(
                "@SQL=(\"urn:schemas:calendar:datereceived\" >= '{0:g}' " +
                "AND \"urn:schemas:mailheader:subject\" LIKE '%{1}%')", DateTime.Today.ToString("g"), wordInSubject);

            string ff1 = "urn:schemas:httpmail:datereceived >= '" + DateTime.Today.ToString("g") + "' AND " + 
                "urn:schemas:mailheader:subject LIKE '%" + wordInSubject + "%' AND " + 
                "urn:schemas:mailheader:sender LIKE '%" + "aaa@bbb.com" + "%'";
            advancedSearch = null;
            Outlook.MAPIFolder folderInbox = null;
            Outlook.MAPIFolder folderSentMail = null;
            Outlook.NameSpace ns = null;
            try {
                ns = OApp.GetNamespace("MAPI");
                folderInbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                folderSentMail = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                scope = "\'" + folderInbox.FolderPath + "\',\'" + folderSentMail.FolderPath + "\'";
                advancedSearch = OApp.AdvancedSearch(scope, ff1, true, advancedSearchTag);
                OApp.AdvancedSearchComplete += Application_AdvancedSearchComplete;
                return advancedSearch;
            }
            catch (Exception ex) {
                Console.WriteLine("\n\tException: " + ex.Message);
                return advancedSearch;
            }
        }
        public void ReadMails() {
            Outlook.Results advancedSearchResults = null;
            Outlook.MailItem resultItem = null;
            StringBuilder strBuilder = null;
            try {
                if (advancedSearch.Tag == advancedSearchTag) {
                    advancedSearchResults = advancedSearch.Results;
                    if (advancedSearchResults.Count > 0) {
                        string[] version = OApp.Version.Split('.');
                        int hostMajorVersion = Convert.ToInt32(version[0]);

                        strBuilder = new StringBuilder();
                        strBuilder.AppendLine("Number of items found: " + advancedSearchResults.Count.ToString());
                        Console.Write("\n [No ] " + "Subject".PadRight(40, '_') + " From".PadRight(30, '_') + " Date".PadRight(30, '_') + "\n\n");
                        for (int i = 1; i <= advancedSearchResults.Count; i++) {
                            resultItem = advancedSearchResults[i] as Outlook.MailItem;
                            if (resultItem != null) {
                                Console.WriteLine(
                                    " [" + i.ToString().PadRight(3) + "] " +
                                    resultItem.Subject.PadRight(41, ' ') +
                                    resultItem.Sender.Name.PadRight(30, ' ') +
                                    resultItem.CreationTime.ToString().PadRight(30, ' '));
                                Marshal.ReleaseComObject(resultItem);
                            }
                        }
                    }
                    else {
                        Console.WriteLine("There are no items found.");
                    }
                }
            }
            catch (Exception ex) {
                Console.WriteLine("\n\tException: " + ex.Message);
            }

            if (resultItem != null) Marshal.ReleaseComObject(resultItem);
            if (advancedSearch != null) Marshal.ReleaseComObject(advancedSearch);
            if (advancedSearchResults != null) Marshal.ReleaseComObject(advancedSearchResults);
        }
        public void Application_AdvancedSearchComplete(Outlook.Search SearchObject) {
            isFinished = true;
        }
    }
}
