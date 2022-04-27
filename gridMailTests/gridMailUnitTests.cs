using Microsoft.VisualStudio.TestTools.UnitTesting;
using gridMail;

namespace gridMailTests
{
    [TestClass]
    public class EmailTests
    {
        [TestMethod]
        public void Create_Simple_Email()
        {
            // Setup test
            string to = "test@test.com";
            string subject = "Test Subject";
            string body = "Body of my email";

            EmailWrapper emailWrapper = new EmailWrapper();
            emailWrapper.subject = subject;
            emailWrapper.body = body;
            emailWrapper.AddRecipient(to);

        }
    }
}