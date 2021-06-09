using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using TestProject1.Base;
using TestProject1.Utilities;

namespace TestProject1.PageObjects
{
    public class Login 
    {
        public IWebDriver driver;

        #region Construtors
        public Login( IWebDriver _driver) 
        {
            driver = _driver;
        }

        #endregion 

        #region Locators

        private string UserNameEdit = "//input";
        private string PasswordEdit = "//input";
        private string LoginBtn = "";




        #endregion
        
        #region Actions 

        public void EnterCredentials(string UserName, string Password)
        {

            driver.FindElement(UserNameEdit, 20).SendKeys(UserName);
            driver.FindElement(PasswordEdit, 20).SendKeys(UserName);
            driver.FindElement(LoginBtn, 20).Click();

        }


        public void Navigateto(string URL)
        {
            driver.Navigate().GoToUrl(URL);
            Thread.Sleep(5000);
        }


        #endregion


    }
}
