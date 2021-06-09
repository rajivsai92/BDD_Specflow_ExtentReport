using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestProject1.Base
{
   public  class BasePage
    {

        public IWebDriver driver { get; set; }

        public BasePage(IWebDriver page)
        {
            this.driver = page;

        }




    }
}
