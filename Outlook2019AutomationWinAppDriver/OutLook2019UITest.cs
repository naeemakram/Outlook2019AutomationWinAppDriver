using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
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

// This code is written by Naeem Akram Malik and provided free of cost as community support for only learning purposes
// In order to learn automated UI testing using WinAppDriver and C#, please visit the Udemy online course discount coupon link given below
//https://www.udemy.com/course/appium-winappdriver-automation-testing/?couponCode=GITHUBAPRIL2020

        [ClassInitialize]
        public static void Setup(TestContext testContext)
        {
            mTestContext = testContext;
            mOptionsOutlook = new AppiumOptions();
            mOptionsOutlook.AddAdditionalCapability("app", "Outlook");
            mOptionsOutlook.AddAdditionalCapability("ms:waitForAppLaunch", Convert.ToString(Outlook2019AutomationWinAppDriver.Properties.Settings.Default.TimeToWaitAfterLaunch));
            // ms:waitForAppLaunch

            mSessionOutlook = new WindowsDriver<WindowsElement>(new Uri($"http://{Outlook2019AutomationWinAppDriver.Properties.Settings.Default.WinAppDriverIP}:{Outlook2019AutomationWinAppDriver.Properties.Settings.Default.WinAppDriverPort}"), mOptionsOutlook);

            var optionsDesktop = new AppiumOptions();
            optionsDesktop.AddAdditionalCapability("app", "Root");
            mSessionDesktop = new WindowsDriver<WindowsElement>(new Uri($"http://{Outlook2019AutomationWinAppDriver.Properties.Settings.Default.WinAppDriverIP}:{Outlook2019AutomationWinAppDriver.Properties.Settings.Default.WinAppDriverPort}"), optionsDesktop);

            mWaitOutlook = new WebDriverWait(mSessionOutlook, TimeSpan.FromSeconds(Outlook2019AutomationWinAppDriver.Properties.Settings.Default.TimeForControLSpecificWait));

            mSessionDesktop.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(Outlook2019AutomationWinAppDriver.Properties.Settings.Default.TimeForImplicitWait);

                mSessionOutlook.Manage().Window.Maximize();
        }

        [TestMethod]
        public void SearchAndMoveMailWithAttachment()
        {
            var homeTab = mSessionOutlook.FindElementByName("Home");
            mWaitOutlook.Until(x => homeTab.Displayed);
            homeTab.Click();

            var searchIcon = mSessionOutlook.FindElementByName("Submit Search");
            searchIcon.Click();

            var hasAttachmentsButton = mSessionOutlook.FindElementByName("Has Attachments");
            mWaitOutlook.Until(x => hasAttachmentsButton.Displayed);
            hasAttachmentsButton.Click();

            System.Threading.Thread.Sleep(3000);

            var allDataItems = mSessionOutlook.FindElementsByTagName("DataItem");
            Debug.WriteLine($"***** Total data items: {allDataItems.Count}");

            int i = 0;

            string subjectToLookFor = "Subject Complicated option";
            string destinationFolderName = "01Udemy";

            WindowsElement mailItem = null;
            string mailName = string.Empty;

            foreach(var mail in allDataItems)
            {
                mailName = mail.GetAttribute("Name");


                Debug.WriteLine($"*****{mailName}");
                
                if(mail.Displayed)
                {
                    //mail.Click();
                    if(mailName.Contains(subjectToLookFor))
                    {
                        mailItem = mail;
                        break;
                    }
                    if(i++ > 10)
                    {
                        break;// prevent very long searches
                    }
                }
            }


            WindowsElement targetFolder = null;

            if(mailItem != null)
            {
                var allTreeNodes = mSessionOutlook.FindElementsByTagName("TreeItem");
                Debug.WriteLine($"Tree nodes found: {allTreeNodes.Count}");

                foreach(var t in allTreeNodes)
                {
                    Debug.WriteLine($"***** {t.GetAttribute("Name")}");
                    if(t.GetAttribute("Name").Contains(destinationFolderName))
                    {
                        targetFolder = t;
                        Debug.WriteLine($"Target folder found {targetFolder.ToString()}");
                        break;
                    }
                }

                if(targetFolder != null)
                {

                    Actions actDrag = new Actions(mSessionOutlook);

                    int offsetX = 0, offsetY = 0;

                    offsetX = targetFolder.Rect.X - mailItem.Rect.X + 5;

                    offsetY = targetFolder.Rect.Y - mailItem.Rect.Y;
                   

                    if (offsetY < 0)// if target folder is above mail item
                    {
                        offsetY-= (targetFolder.Rect.Height / 2);
                    }
                    else // if target folder is below mail item
                    {
                        offsetY += (targetFolder.Rect.Height / 2);
                    }


                    Debug.WriteLine($"Mail item X: {mailItem.Rect.X}, Y: {mailItem.Rect.Y}");
                    Debug.WriteLine($"Target folder X: {targetFolder.Rect.X}, Y: {targetFolder.Rect.Y}");
                    Debug.Write($"Offset X: {offsetX} - X: {offsetY}");                    
                    
                    actDrag.MoveToElement(mailItem, mailItem.Rect.Width / 2, mailItem.Rect.Height / 2);                    
                    actDrag.ClickAndHold(mailItem);
                    actDrag.MoveByOffset(offsetX, offsetY);
                    actDrag.Release(targetFolder);
                    
                    actDrag.Build();
                    actDrag.Perform();
                }
            }

        }


        [TestMethod]
        public void MoveASimpleMail()
        {
            var homeTab = mSessionOutlook.FindElementByName("Home");
            mWaitOutlook.Until(x => homeTab.Displayed);
            homeTab.Click();

            System.Threading.Thread.Sleep(3000);

            var allDataItems = mSessionOutlook.FindElementsByTagName("DataItem");
            Debug.WriteLine($"***** Total data items: {allDataItems.Count}");

            int i = 0;

            string subjectToLookFor = "JazzCashAlert";
            string destinationFolderName = "01Udemy";

            WindowsElement mailItem = null;
            string mailName = string.Empty;

            foreach (var mail in allDataItems)
            {
                mailName = mail.GetAttribute("Name");


                Debug.WriteLine($"*****{mailName}");

                if (mail.Displayed)
                {
                    if (mailName.Contains(subjectToLookFor))
                    {
                        mailItem = mail;
                        break;
                    }
                    if (i++ > 10)
                    {
                        break;// prevent very long searches
                    }
                }
            }


            WindowsElement targetFolder = null;

            if (mailItem != null)
            {
                var allTreeNodes = mSessionOutlook.FindElementsByTagName("TreeItem");
                Debug.WriteLine($"Tree nodes found: {allTreeNodes.Count}");

                foreach (var t in allTreeNodes)
                {
                    Debug.WriteLine($"***** {t.GetAttribute("Name")}");
                    if (t.GetAttribute("Name").Contains(destinationFolderName))
                    {
                        targetFolder = t;
                        Debug.WriteLine($"Target folder found {targetFolder.ToString()}");
                        break;
                    }
                }

                if (targetFolder != null)
                {

                    Actions actDrag = new Actions(mSessionOutlook);

                    int offsetX = 0, offsetY = 0;

                    offsetX = targetFolder.Rect.X - mailItem.Rect.X + 5;

                    offsetY = targetFolder.Rect.Y - mailItem.Rect.Y;


                    if (offsetY < 0)// if target folder is above mail item
                    {
                        offsetY -= (targetFolder.Rect.Height / 2);
                    }
                    else // if target folder is below mail item
                    {
                        offsetY += (targetFolder.Rect.Height / 2);
                    }


                    Debug.WriteLine($"Mail item X: {mailItem.Rect.X}, Y: {mailItem.Rect.Y}");
                    Debug.WriteLine($"Target folder X: {targetFolder.Rect.X}, Y: {targetFolder.Rect.Y}");
                    Debug.Write($"Offset X: {offsetX} - X: {offsetY}");

                    actDrag.MoveToElement(mailItem, mailItem.Rect.Width / 2, mailItem.Rect.Height / 2);
                    actDrag.ClickAndHold(mailItem);
                    actDrag.MoveByOffset(offsetX, offsetY);
                    actDrag.Release(targetFolder);
                    actDrag.Build();
                    actDrag.Perform();
                }
            }

        }

        [TestMethod]
        public void SelectOtherAndSelectUnRead()
        {
            ///Button[@Name=\"Other\"]
            var otherButton = mSessionOutlook.FindElementByXPath("//Button[@Name=\"Other\"]");
            mWaitOutlook.Until(x => otherButton.Displayed);

            otherButton.Click();


            // /Button[@Name=\"Sort, arrange or filter messages\"][@ClassName=\"NetUISimpleButton\"]

            var comboBy = mSessionOutlook.FindElementByXPath("//Button[@Name=\"Sort, arrange or filter messages\"][@ClassName=\"NetUISimpleButton\"]");
            mWaitOutlook.Until(x => comboBy.Displayed);
            comboBy.Click();

            var menuItemUnread = mSessionOutlook.FindElementByXPath("//MenuItem[@Name=\"Unread Mail\"][@ClassName=\"NetUITWBtnCheckMenuItem\"]");
            mWaitOutlook.Until(x => menuItemUnread.Displayed);
            menuItemUnread.Click();

        }

        [TestMethod]
        public void ToToSendReceiveAndUpdateFolder()
        {
            var sendReceiveTab = mSessionOutlook.FindElementByName("Send / Receive");
            mWaitOutlook.Until(x => sendReceiveTab.Displayed);
            sendReceiveTab.Click();

            var btnUpdateFolder = mSessionOutlook.FindElementByXPath("//Button[@Name=\"Update Folder\"][@ClassName=\"NetUIRibbonButton\"]");
            mWaitOutlook.Until(x => btnUpdateFolder.Displayed);
            btnUpdateFolder.Click();
        }

    }
}
