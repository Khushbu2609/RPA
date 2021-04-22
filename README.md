# Group3RPAProject:

Objective:
Project creates new excel sheet and write extracted data for Canada Covid19 from given URL
Once excel file is ready with extracted data it will send email to configured receiptant/s.
It also pops up message boxes with relavant messages for performed action/s.

Tasks:
1. create blank excel sheet at configured location named "canadaCovid19Report.xlsx"
2. extract canada covid-19 data from URL: "https://health-infobase.canada.ca/covid-19/?stat=num&measure=total#a2"
3. write extracted data to the above excel sheet
4. send Email with excel file as attachement to configured receiver

Configure following before running project:
1. configure location for the excel sheet:
	go to project->main.xaml->double click inside first "Excel Application Scope"
	edit the file location with full path according to your suitability
	click 'Flowchart' on top of the window to go back 
2. Do the same for second "Excel Applicaton Scope"
3. Add your Email address to recieve email with excel as attachment
	go to project->main.xaml->double click inside "Send SMTP Mail Message"
	click on Properties window on the right
	edit "Receiver" field, you can put your email address in "To"(mandatory), "Bcc" or "cc"
4. Once above changes are done you should be able to click "Debug File" -> "Run File" to see how its working
   click "ok" on message boxes every time to continue.
