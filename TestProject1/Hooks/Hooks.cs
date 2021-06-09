using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TechTalk.SpecFlow;
using TestProject1.Utilities;
using BoDi;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports.Gherkin.Model;
using System.Reflection;
using System.IO;
using System.Threading;
using TechTalk.SpecFlow.Infrastructure;
using System.Diagnostics;
using NUnit.Framework;

namespace TestProject1.Hooks
{
    [Binding]
    public sealed class Hooks
    {

        private IObjectContainer _IObjectContainer;
        private Browser _browser;
        private static ExtentTest featureName;
        private static ExtentTest scenario;
        private static ExtentReports extent;
        public static string ScreenshotPath;
        public IWebDriver driver { get; set; }
        private static ScenarioContext _scenarioContext;
        public static string reportPath = CommonFuntions.GetProjectPathOfFolder("TestReports");
        public static string SubFolderName;



        public Hooks(IObjectContainer objectContainer, Browser browser, ScenarioContext scenarioContext)
        {
            _IObjectContainer = objectContainer;
            _browser = browser;
            _scenarioContext = scenarioContext;


        }



        [BeforeTestRun]
        public static void InitializeReport()
        {

            SubFolderName = "AutomationResults" + DateTime.Today.ToString("ddMMyyyy") + "_" + DateTime.Now.ToShortTimeString().Replace(':', '_').Replace(" ", "_");
            Directory.CreateDirectory(reportPath + @"\" + SubFolderName + @"\");
            string fileName = reportPath + @"\" + SubFolderName + @"\" + "AutomationResult" + "_" + DateTime.Now.ToString("ddMMyyyy_hh-mm-ss") + ".html";
            var htmlReporter = new ExtentHtmlReporter(fileName);
            htmlReporter.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Standard;
            htmlReporter.Config.DocumentTitle = "Automation Execution Report";

            extent = new ExtentReports();
            extent.AttachReporter(htmlReporter);


        }


        [BeforeFeature]
        public static void BeforeFeature(FeatureContext featureContext)
        {

            featureName = extent.CreateTest<Feature>(featureContext.FeatureInfo.Title);

        }



        [BeforeScenario]
        public void BeforeScenario()
        {
            if (driver == null)
            {
                driver = _browser.GetBrowser();

            }

            _IObjectContainer.RegisterInstanceAs(driver);
            scenario = featureName.CreateNode<Scenario>(_scenarioContext.ScenarioInfo.Title);

        }


        [AfterStep]
        public void InsertReportingSteps()
        {
            #region Log to Extent Report

            var stepType = _scenarioContext.StepContext.StepInfo.StepDefinitionType.ToString();


            if (_scenarioContext.ScenarioExecutionStatus.ToString() == "StepDefinitionPending")
            {

                if (stepType == "Given")
                    scenario.CreateNode<Given>(_scenarioContext.StepContext.StepInfo.Text).Skip("Step Definition Pending");

                else if (stepType == "When")
                    scenario.CreateNode<When>(_scenarioContext.StepContext.StepInfo.Text).Skip("Step Definition Pending");

                else if (stepType == "Then")
                    scenario.CreateNode<Then>(_scenarioContext.StepContext.StepInfo.Text).Skip("Step Definition Pending");

                else if (stepType == "And")
                    scenario.CreateNode<And>(_scenarioContext.StepContext.StepInfo.Text).Skip("Step Definition Pending");

                else if (stepType == "But")
                    scenario.CreateNode<But>(_scenarioContext.StepContext.StepInfo.Text).Skip("Step Definition Pending");


            }

            else if (_scenarioContext.TestError == null)
            {
                if (stepType == "Given")
                    scenario.CreateNode<Given>(_scenarioContext.StepContext.StepInfo.Text);

                else if (stepType == "When")
                    scenario.CreateNode<When>(_scenarioContext.StepContext.StepInfo.Text);

                else if (stepType == "Then")
                    scenario.CreateNode<Then>(_scenarioContext.StepContext.StepInfo.Text);

                else if (stepType == "And")
                    scenario.CreateNode<And>(_scenarioContext.StepContext.StepInfo.Text);

                else if (stepType == "But")
                    scenario.CreateNode<But>(_scenarioContext.StepContext.StepInfo.Text);
            }

            else if (_scenarioContext.TestError != null)
            {

                #region Error Message 
                string Failmsg = "<br><b>Message:</b> " + _scenarioContext.TestError.Message +
                                "<br><b>Inner Exception: </b> " + _scenarioContext.TestError.InnerException +
                                 "<br><b>Stack Trace: </b>" + _scenarioContext.TestError.StackTrace;
                #endregion

                #region ScreenShot
                MediaEntityModelProvider MMP = MediaEntityBuilder.CreateScreenCaptureFromBase64String
                    (((ITakesScreenshot)driver).GetScreenshot().AsBase64EncodedString, "Image" + DateTime.Now).Build();
                #endregion               

                #region Write to Extent reports
                if (stepType == "Given")
                    scenario.CreateNode<Given>(_scenarioContext.StepContext.StepInfo.Text).Fail(Failmsg, MMP);

                else if (stepType == "When")
                    scenario.CreateNode<When>(_scenarioContext.StepContext.StepInfo.Text).Fail(Failmsg, MMP);

                else if (stepType == "Then")
                    scenario.CreateNode<Then>(_scenarioContext.StepContext.StepInfo.Text).Fail(Failmsg, MMP);

                else if (stepType == "And")
                    scenario.CreateNode<And>(_scenarioContext.StepContext.StepInfo.Text).Fail(Failmsg, MMP);

                else if (stepType == "But")
                    scenario.CreateNode<But>(_scenarioContext.StepContext.StepInfo.Text).Fail(Failmsg, MMP);
                #endregion
            }
            #endregion 

        }

        [AfterScenario]
        public void AfterScenario()
        {
            //TODO: implement logic that has to run after executing each scenario

            driver.Close(); driver.Quit();
            _browser.driver = null;
        }

        [AfterTestRun]
        public static void AfterTestRun()
        {

            extent.Flush();

        }
    }
}


