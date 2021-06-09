using System;
using TechTalk.SpecFlow;
using NUnit.Framework;



namespace TestProject1.FeatureSteps
{
    [Binding]
    public class MyAccountsSteps
    {
        [Given(@"Verify Top Navigation")]
        public void GivenVerifyTopNavigation()
        {
            //ScenarioContext.Current.Pending();
        }
        
        [Given(@"Click on Account Link")]
        public void GivenClickOnAccountLink()
        {
            //ScenarioContext.Current.Pending();
        }
        [Then(@"Verify Home page(.*)")]
        public void ThenVerifyHomePage(int p0)
        {
            // ScenarioContext.Current.Pending();
            Assert.Fail("Failed");
        }
        [Given(@"Enter (.*) and (.*) and Click on Submit")]
        public void GivenEnterAndAndClickOnSubmit(int p0, int p1)
        {
           // ScenarioContext.Current.Pending();
        }


        [Then(@"Verify account page")]
        public void ThenVerifyAccountPage()
        {
           // ScenarioContext.Current.Pending();
        }

     
    }
}
