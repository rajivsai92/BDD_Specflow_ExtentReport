Feature:3_MyAccounts
	Simple calculator for adding two numbers


@mytag

Scenario: (1) Home page Functionality
	Given Launch Application with URL
	| URL |
	| https://www.youtube.com/|
	And   Enter Username and Password and Click on Submit
	| Username | Password |
	| rajiv    | test     |
	
Scenario:(2) Home page Functionality2
	Given Launch Application with URL
	| URL |
	| https://www.youtube.com/ |
	And   Enter Username and Password and Click on Submit
	| Username | Password |
	| sai    | test2     |
#Scenario Outline: Account Page
#	Given Launch Application
#	And   Enter <Username> and <Password> and Click on Submit
#	Then  Verify Home page1
#Examples: 
#| Username | Password | 
#| 50   | 70    |
#| 60    | 80     |


