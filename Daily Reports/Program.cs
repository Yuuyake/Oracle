using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;

namespace Daily_Reports {
    class Program {
        static void Main(string[] args) {
            Console.Title = "OUTLOOK SEARCH";
            var src = new Searcher();

            src.RunAdvancedSearch("IMPERVA DATABASE");
            Console.Write("\n Searching mails of \"IMPERVA DATABASE\"");

            int counter = 1;
            while (true) {
                Console.Write("\r Searching mails of \"IMPERVA DATABASE\"" + " ".PadRight(counter%5,'.') + " ".PadRight(5- counter%5));
                if (src.isFinished == true) {
                    Console.Write("\r Search is DONE. Reading mails . . .                               \n\n");
                    src.ReadMails();
                    break;
                }
                else {
                    Thread.Sleep(500);
                    counter++;
                }
            }
            Console.ReadLine();
        }
    }
}
