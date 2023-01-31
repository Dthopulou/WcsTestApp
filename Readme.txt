1. Command - SCAN-RECOVERY-FOLDER

This command scans 'Recoverable folder' with mapped folder or mailbox sync search criteria for deleted items based on the parameter

Syntax: <Command> <ImpersonatorSMTP> <password> <exchange server> <ExchangeVersion> <SearchFor> <IncludePRPolicyTag> <CountOnly> <StartDate> <EndDate> <ReportMode> <ExtractEmail> <ExtractCalendar> <ExtractContact> <ExtractTask> <ExtractNote> <ExtractMiscellaneous>");

Example: SCAN-RECOVERY-FOLDER ewsuser@exdev2016.local !manage6 10.192.211.238 Exchange2010_SP2 2 True FALSE 2018-04-19 2018-04-23 FALSE False True False True True False

Parameter Details
	<exchangeServer>: If this has clustername then this needs to be in quotes. Example "<ClusterName>ExchangeServer"
	
	ExchangeVersion: Parameter should be one of these
		Exchange2010_SP2
		Exchange2010	
		Exchange2010_SP1	 
		Exchange2010_SP2
		Exchange2013
		Exchange2013_SP1
		Exchange2016

	SearchFor: 
		1: WCSE_FolderMappings 
		2: WCSE_SFMailboxSync

	PolicyTags: 
		True\False
		


2. Users.txt 
	This file should have all the emails addresses
	Each line should have one email address
	Should be located at EWSTestApp.exe directory

4. Microsoft.Exchange.WebServices.dll must be located at EWSTestApp.exe directory

5. Runs only in report mode

6. <ExtractEmail> : True Extracts only Email
7. <ExtractCalendar> : True Extracts only Calendar
8. <ExtractContact> : True - Extracts Contact
9. <ExtractTask> : True - Extracts Task
10. <ExtractNote> : True - Extracts Sticky Note
11. <ExtractMiscellaneous> : True - if anything other than above mentioned items will be extracted

NOTE: this command still creates folder under inbox but Calendar, Contact, Task, Note items will be restored to respective outlook folders.
There are 17 parameters, take time when you are constructing the parameters



