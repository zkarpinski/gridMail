using System;
using System.Collections.Generic;
using System.IO;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;


namespace gridMail {
    public class EmailWrapper {
        // Public properties
        public string body { get; set; }
        public string subject { get; set; }
        public string sender { get; set; }


        // Declare the variables
        private bool use_html = false;
        private string template_file;
        private string html_body;
        private List<string> attachments;
        private List<Tuple<string, string>> inline_attachments;
        private List<string> recipients;
        private List<string> distribution_groups;
        private ExchangeService exchange_service;
        private string exchange_uri = ConfigurationManager.AppSettings.Get("exchange_uri");



        // Constructor
        public EmailWrapper() {
            this.attachments = new List<string>();
            this.inline_attachments = new List<Tuple<string, string>>();
            this.recipients = new List<string>();
            this.distribution_groups = new List<string>();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns>Integer 1: Success, -1 = Fail</returns>
        public int Send() {
            // Connect to exchange_service
            exchange_service = new ExchangeService(ExchangeVersion.Exchange2010);
            exchange_service.Url = new Uri(exchange_uri);
            exchange_service.UseDefaultCredentials = true;

            // Create the email
            EmailMessage emailMessage = new EmailMessage(exchange_service);

            // Subject
            emailMessage.Subject = this.subject;

            //Body
            if (use_html) {
                emailMessage.Body = new MessageBody(BodyType.HTML, this.html_body);

            }
            else { 
                emailMessage.Body = new MessageBody(BodyType.HTML, this.body);
            }

            // Add Recipients
            foreach (string rep in this.recipients) {
                emailMessage.ToRecipients.Add(rep);
            }
            // Add Distribution List Persons
            foreach (string dist in this.distribution_groups) {
                ExpandGroupResults grp = exchange_service.ExpandGroup(dist);
                foreach (EmailAddress em in grp.Members) {
                    emailMessage.ToRecipients.Add(em.Address);
                }
            }


            //Add Attachements
            foreach (string att in this.attachments) {
                emailMessage.Attachments.AddFileAttachment(att);
            }

            // Add Inline attachments
            //https://msdn.microsoft.com/en-us/library/office/hh532564(v=exchg.80).aspx
            foreach (Tuple<string,string> tup in this.inline_attachments) {
                string contentID = tup.Item1;
                string file = tup.Item2;

                FileAttachment fileAttachment = emailMessage.Attachments.AddFileAttachment(file);
                fileAttachment.ContentId = contentID;
                fileAttachment.IsInline = true;
            }


            // Save and send
            try {
                emailMessage.Save();
                emailMessage.SendAndSaveCopy();   
            }
            catch {
                return -1;
            }
            return 0;
        }

        public void AddAttachment(string file) {
            //Verify file exists
            if (File.Exists(file)) {
                this.attachments.Add(file);
            }
            else {
                // Give error message
                Console.WriteLine("{0} the file does not exist.\nPress Enter to exit...", file);
                Console.ReadLine();
                Environment.Exit(-1);
            }
        }

        public void AddInlineAttachment(string s) {
            // Split the text.
            string cID, file;
            cID = "<" + s.Split(';')[0] + ">";
            file = s.Split(';')[1];
            Tuple<string, string> tup = new Tuple<string, string>(cID, file);

            //Verify file exists
            if (File.Exists(file)) {
                this.inline_attachments.Add(tup);
            }
            else {
                // Give error message
                Console.WriteLine("{0} the file does not exist.\nPress Enter to exit...", file);
                Console.ReadLine();
                Environment.Exit(-1);
            }
        }

        public void AddRecipient(string address) {
            this.recipients.Add(address);
        }

        public void AddGroup(string grp) {
            this.distribution_groups.Add(grp);
        }

        public void GetExchangeURI()
        {
            Console.WriteLine("The exchange uri is: {0} the file does not exist.\nPress Enter to exit...", this.exchange_uri);
        }

        public void UseHTML(string file) {
            //Verify file exists
            if (File.Exists(file)) {
                this.template_file = file;
                this.html_body = File.ReadAllText(file);
                this.use_html = true;
            }
            else {
                // Give error message
                Console.WriteLine("{0} the file does not exist.\nPress Enter to exit...", file);
                Console.ReadLine();
                Environment.Exit(-1);
            }
        }
    }
}
