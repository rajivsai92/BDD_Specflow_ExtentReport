using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Edge.SeleniumTools;

namespace TestProject1.Utilities
{
    public class Browser
    {
        public  IWebDriver driver { get; set; }


        public  IWebDriver GetBrowser()
        {


            
            var optionsEdge = new EdgeOptions();
            optionsEdge.UseChromium = true;

            driver = new EdgeDriver(@"C:\Users\rajivsa_90\source\repos\TestProject1\TestProject1\Driver", optionsEdge);
            driver.Navigate().GoToUrl("https://www.google.com/");
            //driver.Manage().Window.Maximize();
            return driver;


        }
             
    }
}
