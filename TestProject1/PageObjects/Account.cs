using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestProject1.PageObjects
{
    public class Account
    {
        public IWebDriver driver;

        #region Construtors
        public Account(IWebDriver _driver)
        {
            driver = _driver;
        }

        #endregion

        public void EnterAccountDetails()
        {

           

        }

    }
}
