// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Security;

namespace EasyReproExtentReports
{
    [TestClass]
    public class Custom_Funktion : TestBase
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"]);

        [TestMethod]
        [Description("OpenAccount")]
        public void UCITestOpenActiveAccount()
        {

            try
            {
                var client = new WebClient(TestSettings.Options);
                using (var xrmApp = new XrmApp(client))
                {
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                    AddScreenShot(client, "After login");

                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    AddScreenShot(client, "After OpenSubArea: Sales/Accounts");

                    xrmApp.Grid.Search("Adventure");
                    AddScreenShot(client, "After Search: Adventure");

                    xrmApp.Grid.OpenRecord(0);

                    xrmApp.ThinkTime(3000);

                    xrmApp.Entity.Save();
                    AddScreenShot(client, "After Save");

                    Assert.IsTrue(true);

                }

            }
            catch (Exception e)
            {
                LogExceptionAndFail(e);
            }
         }

        [TestMethod]
        [Description("Test should fail with no error")]
        public void JustFail()
        {
            // Example log entries
            Test.Info("Log an information message");
            Test.Warning("Log a warning message");

            // Test code here

            Assert.Fail("Test failed because the passing criteria was not met");
        }


        [TestMethod]
        public void UCITestGetActiveGridItems()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.Grid.GetGridItems();

                xrmApp.Grid.Sort("Account Name");

                xrmApp.ThinkTime(3000);
            }
        }

        [TestMethod]
        public void UCITestOpenTabDetails()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.Grid.SwitchView("All Accounts");

                xrmApp.Grid.OpenRecord(0);

                xrmApp.ThinkTime(3000);

                xrmApp.Entity.SelectTab("Details");

                xrmApp.Entity.SelectTab("Related", "Contacts");

                xrmApp.ThinkTime(3000);
            }
        }

        [TestMethod]
        public void UCITestGetObjectId()
        {
            var client = new WebClient(TestSettings.Options);

            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.Grid.OpenRecord(0);

                Guid objectId = xrmApp.Entity.GetObjectId();

                xrmApp.ThinkTime(3000);
            }
        }

        [TestMethod]
        public void UCITestOpenSubGridRecord()
        {
            var client = new WebClient(TestSettings.Options);

            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.Grid.OpenRecord(0);

                xrmApp.Entity.GetSubGridItems("CONTACTS");

                xrmApp.ThinkTime(3000);
            }
        }
    }
}