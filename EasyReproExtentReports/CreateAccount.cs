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
    public class CreateAccount : TestBase
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"]);

        [TestMethod]
        [Description("Test should pass")]
        public void CreateAccount_Pass()
        {
            try
            {
                var client = new WebClient(TestSettings.Options);
                using (var xrmApp = new XrmApp(client))
                {
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                    AddScreenShot(client, "After login");

                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    AddScreenShot(client, $"After OpenApp: {UCIAppName.Sales}");

                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    AddScreenShot(client, "After OpenSubArea: Sales/Accounts");

                    xrmApp.CommandBar.ClickCommand("New");
                    AddScreenShot(client, "After ClickCommand: New");

                    xrmApp.Entity.SetValue("name", TestSettings.GetRandomString(5, 15));
                    AddScreenShot(client, "After SetValue: name");

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
            // Test code here

            Assert.Fail("Test failed because the passing criteria was not met");
        }

        [TestMethod]
        [Description("Test should fail due to an error")]
        public void CreateAccount_Error()
        {
            try
            {
                var client = new WebClient(TestSettings.Options);
                using (var xrmApp = new XrmApp(client))
                {
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                    AddScreenShot(client, "After login");

                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                    AddScreenShot(client, $"After OpenApp: {UCIAppName.Sales}");

                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    AddScreenShot(client, "After OpenSubArea: Sales/Accounts");

                    xrmApp.CommandBar.ClickCommand("New");
                    AddScreenShot(client, "After ClickCommand: New");

                    // Field name is incorrect which will cause an exception
                    xrmApp.Entity.SetValue("name344543", TestSettings.GetRandomString(5, 15));
                    AddScreenShot(client, "After SetValue: name");

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
    }
}