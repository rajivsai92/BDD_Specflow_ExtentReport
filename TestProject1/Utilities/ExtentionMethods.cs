using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestProject1.Utilities
{
    public static class ExtentionMethods
    {

        public static IWebElement FindElement(this IWebDriver driver, string Locator, int timeoutInSeconds)
        {

            By by = CommonFuntions.GetBy(Locator);

            if (timeoutInSeconds > 0)
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
                return wait.Until(drv => drv.FindElement(by));
            }

            return driver.FindElement(by);


        }

        public static IList<IWebElement> FindElements(this IWebDriver driver, By by, int timeoutInSeconds)
        {
            if (timeoutInSeconds > 0)
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
                return wait.Until(drv => (drv.FindElements(by).Count > 0) ? drv.FindElements(by) : null);

            }
            return driver.FindElements(by);
        }

        public static bool WaitUntillTitleContains(this IWebDriver driver, string Title, int timeoutInSeconds)
        {
            bool flag = false;
            if (timeoutInSeconds > 0)
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.TitleContains(Title));
                flag = true;
            }

            return flag;


        }

    }
}

