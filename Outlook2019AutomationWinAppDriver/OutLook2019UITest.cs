using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;

namespace Outlook2019AutomationWinAppDriver
{
    [TestClass]
    public class OutLook2019UITest
    {
        static TestContext mTestContext;
        static WindowsDriver<WindowsElement> mSessionOutlook;
        static WindowsDriver<WindowsElement> mSessionDesktop;
        static AppiumOptions mOptionsOutlook;

        static WebDriverWait mWaitOutlook;

        [ClassInitialize]
        public static void Setup(TestContext testContext)
        {
            mTestContext = testContext;
            mOptionsOutlook = new AppiumOptions();
            mOptionsOutlook.AddAdditionalCapability("app", "Outlook");
            mOptionsOutlook.AddAdditionalCapability("ms:waitForAppLaunch", "15");
            // ms:waitForAppLaunch

            mSessionOutlook = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), mOptionsOutlook);

            var optionsDesktop = new AppiumOptions();
            optionsDesktop.AddAdditionalCapability("app", "Root");
            mSessionDesktop = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), optionsDesktop);

            mWaitOutlook = new WebDriverWait(mSessionOutlook, TimeSpan.FromSeconds(5));

            mSessionOutlook.Manage().Window.Maximize();
        }

        [TestMethod]
        public void OpenOutLook()
        {
            var homeTab = mSessionOutlook.FindElementByName("Home");
            mWaitOutlook.Until(x => homeTab.Displayed);
            homeTab.Click();

            var searchIcon = mSessionOutlook.FindElementByName("Submit Search");
            searchIcon.Click();

            var hasAttachmentsButton = mSessionOutlook.FindElementByName("Has Attachments");
            mWaitOutlook.Until(x => hasAttachmentsButton.Displayed);
            hasAttachmentsButton.Click();

            //var messagesTable = mSessionOutlook.FindElementByName("Table View");
            //messagesTable.Click();

            //var allDataItems = messagesTable.FindElementsByTagName("DataItem");
            //Debug.WriteLine($"***** Total data items: {allDataItems.Count}");
            //foreach(var mail in allDataItems)
            //{
            //    Debug.WriteLine($"***** {mail.Text} \r\n {mail.GetAttribute("Name")}");
            //}
        }
    }
}
