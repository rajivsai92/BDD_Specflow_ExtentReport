using NUnit.Framework;
using System;
using TechTalk.SpecFlow;
using TechTalk.SpecFlow.Assist;
using TestProject1.PageObjects;


namespace TestProject1.FeatureSteps
{
    [Binding]
    public class LoginFeatureSteps
    {
        private Login _Login;
        private Account _act;
        private ScenarioContext _scenarioContext;
   
        public LoginFeatureSteps(Login login, Account act,  ScenarioContext scenarioContext)
        {

            _Login = login;
            _act = act;
            _scenarioContext = scenarioContext;

        }

        [Given(@"Launch Application with URL")]
        public void GivenLaunchApplicationWithURL(Table table)
        {

            dynamic links = table.CreateDynamicInstance();
            _Login.Navigateto(links.URL);

        }


        [Given(@"Launch Application")]
        public void GivenLaunchApplication()
        {
            
        }

        [Given(@"Enter Username and Password and Click on Submit")]
        public void GivenEnterUsernameAndPasswordAndClickOnSubmit(Table table)
        {
            //ScenarioContext.Current.Pending();
        }

       


        [Then(@"Verify Home page")]
        public void ThenVerifyHomePage()
        {
            //_scenarioContext.Pending();
            //Console.WriteLine("a");
            Assert.Fail("Intentionally failed");


        }


       


    }
}
