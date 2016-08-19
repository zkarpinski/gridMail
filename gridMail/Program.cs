using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace gridMail {
    class Program {
        static void Main(string[] args) {

            // Instantiate the email wrapper
            EmailWrapper new_email = new EmailWrapper();

            // Parse the args and pass the email wrapper
            ParseArgs(args, ref new_email);

            // Send Email
            int result = new_email.Send();

            Environment.Exit(result);
        }

        private static void PrintHelpMenu() {
            Console.WriteLine("\nusage: gridMail.exe [-a -b -f -g -h -i -s -t -w]");
            Console.WriteLine("\t-a --attachment [item1,item2]");
            Console.WriteLine("\t-b --body [Email body text]");
            Console.WriteLine("\t-g --group [group@domain.com,group2@domain.com]");
            Console.WriteLine("\t-i --inplace [contentID;file,contentId;file2]");
            Console.WriteLine("\t-f --from [sender@domain.com]");
            Console.WriteLine("\t-t --to [person@domain.com,person2@domain.com]");
            Console.WriteLine("\t-w --web [path_to_html_file]   Use HTML template");
            Console.WriteLine(@"example: gridMail.exe -a ""item1.txt,item2.txt"" -t ""person1@domain.com,person2@email.com"" -b ""Body goes here"" -s ""Subject Line"" -f sender@email.com");
            Console.WriteLine("\nPress enter to exit...");
            Console.ReadLine();
            Environment.Exit(0);
        }

        private static void ParseArgs(string[] args, ref EmailWrapper email) {
            //Read each argument
            int args_count = args.Length;

            if (args_count <= 0) { PrintHelpMenu(); }
            for (int i = 0; i < args_count; i++) {

                switch (args[i]) {
                    case "-h":
                    case "--help":
                        PrintHelpMenu();
                        break;
                    case "-a": //Attachments
                    case "--attachment":
                        // Split the next argument by semicolons and commas
                        if (i + 1 < args_count) {
                            foreach (string att in args[i + 1].Split(new Char[] { ',', ';' })) {
                                email.AddAttachment(att);
                            }
                            i++;
                        }
                        else {
                            Console.WriteLine("No argument proceeding the attachments list [-a]");
                            PrintHelpMenu();
                        }
                        break;
                    case "-b": // Email body
                    case "--body":
                        if (i + 1 < args_count) {
                            email.body = args[i + 1];
                            i++;
                        }
                        else {
                            Console.WriteLine("No argument proceeding the body argument [-b]");
                            PrintHelpMenu();
                        }
                        break;
                    case "-g": // Distribution Groups
                    case "--group":
                        // Split the next argument by semicolons and commas
                        if (i + 1 < args_count) {
                            foreach (string grp in args[i + 1].Split(new Char[] { ',', ';' })) {
                                email.AddGroup(grp);
                            }
                            i++;
                        }
                        else {
                            Console.WriteLine("No argument proceeding the distrubution group list [-g]");
                            PrintHelpMenu();
                        }
                        break;
                    case "-i": //Inline Attachments
                    case "--inline":
                        // Split the next argument by semicolons and commas
                        if (i + 1 < args_count) {
                            foreach (string s in args[i + 1].Split(',')) {
                                email.AddInlineAttachment(s);
                            }
                            i++;
                        }
                        else {
                            Console.WriteLine("No argument proceeding the attachments list [-a]");
                            PrintHelpMenu();
                        }
                        //Console.WriteLine("Inline attachments not yet supported");
                        i++;
                        break;
                    case "-t": // Recipients
                    case "--to":
                        if (i + 1 < args_count) {
                            foreach (string t in args[i + 1].Split(new Char[] { ',', ';' })) {
                                email.AddRecipient(t);
                            }
                            i++;
                        }
                        else {
                            Console.WriteLine("No argument proceeding the recipients list [-t]");
                            PrintHelpMenu();
                        }
                        break;
                    case "-s": //Subject Line
                    case "--subject":
                        if (i + 1 < args_count) {
                            email.subject = args[i + 1];
                            i++;
                        }
                        else {
                            Console.WriteLine("No argument proceeding the subject argument [-s]");
                            PrintHelpMenu();
                        }
                        break;
                    case "-f": // Sender
                    case "--from":
                        if (i + 1 < args_count) {
                            email.sender = args[i + 1];
                            i++;
                        }
                        else {
                            Console.WriteLine("No argument proceeding the from argument [-f]");
                            PrintHelpMenu();
                        }
                        break;
                    case "-w": // Use HTML Template
                    case "--web":
                        if (i + 1 < args_count) {
                            email.UseHTML(args[i + 1]);
                            i++;
                        }
                        else {
                            Console.WriteLine("No argument proceeding the from argument [-w]");
                            PrintHelpMenu();
                        }
                        break;
                    default:
                        Console.WriteLine("Invaild input: {0}", args[i]);
                        PrintHelpMenu();
                        break;
                }
            }
        }
    }
}
