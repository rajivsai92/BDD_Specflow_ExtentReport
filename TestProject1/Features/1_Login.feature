Feature:1_LoginFeature
	Login To MID Application

@mytag
Scenario: (1) Home page Functionality
	Given Launch Application with URL
	| URL |
	| https://www.facebook.com/ |
	And   Enter Username and Password and Click on Submit
	| Username | Password |
	| rajiv    | test     |
	
Scenario: (2) Home page Functionality2
	Given Launch Application with URL
	| URL |
	| https://www.facebook.com/ |
	And   Enter Username and Password and Click on Submit
	| Username | Password |
	| sai    | test2     |
	Then  Verify Home page
