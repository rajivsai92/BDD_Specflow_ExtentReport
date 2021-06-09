using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestProject1.Utilities
{
    public class CommonFuntions
    {
        public static By GetBy(string locator)
        {

            string[] locators = locator.Split('>', 2);

            By by = null;
            switch (locators[0])
            {
                case "Xpath": by = By.XPath(locators[1]); break;
                case "ID": by = By.Id(locators[1]); break;
                case "Class": by = By.ClassName(locators[1]); break;
                case "TagName": by = By.TagName(locators[1]); break;
                case "Name": by = By.Name(locators[1]); break;
                default: throw new Exception("Invalid locator");

            }

            return by;

        }


        public static string GetProjectPathOfFolder(string folderName)
        {
            string path = string.Empty;
            try
            {

                string projectDir = (new FileInfo(AppContext.BaseDirectory)).Directory.Parent.Parent.Parent.FullName;
                path = projectDir + @"\" + folderName + @"\";

         
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return path;
        }

    }
}
