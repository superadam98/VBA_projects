# VBA_projects

Project that will be focues on creating a nested project management tool within excel. Can be easily imported to servers for colaborative work.

Current Project Readme and settings
1) To create a new spreadsheet, press the create new spreadsheet button that calls the main copy macro
2) references I have enabled are Visual Basic For Applications, Microsoft Excel 16.0 Object Library, OLE Automation, Microsoft Office 16.0 Object Library, Microsoft Forms 2.0 Object Library, Microsoft Outlook 16.0 Object Library
3) Disabled Ask to update autopmatic links, Cause that POP UP IS PAIN WHEN IM TESTING lol (under advanced and general incase you dont know)
4) Enabled vba macros in the trust center, and allowed external content after a prompt, despite step 3

DA RULES
1) Ensure inputs are correct, incorrect inputs can ruin data, follow prompts and go slow
2) Do not change the Project Title or subtitles once created, this includes the excel sheet name 
3) Always accept anything that excel asks before using this chart
4) If a new project needs to be added to an exisitng project list CONTACT IT, this is difficult to accomplish without breaking the other charts(Temporary rule)
5) You can change the number of subprojects under a project by moving the other projects down and copying the enitre row.
6) Do not create a new sheet that already has a hyper link
7) Only change values for subprojects WITHOUT Hyper Links
8) Write notes and other information on a different sheet. They will be saved.
9) Go to File/Options/Advanced/General, and uncheck the checkbox that says "Ask to update automatic links"
10) Go to File/Options/TrustCenter and enable all data updates(not recommended), but do so at your own peril, this will decrease pop ups when using this, but leaves you vulnerable to downloading excel sheets with bad macros.

Planned updates: 
1) Cleaning up code and unused variables
2) Writing notes for code
3) adding a function during the save process that can add lines after a child gantt chart is created
4) adding a function that will load new lines to a parent if a new subproject is created on a child
5) add to functions so the actual cost field can get automatically updated as it is requested
6) add the ability to approve certain budget requirements
7) send email from the send email button, will launch via outlook
8) potentially have start or end dates interact with the (today)names space so gantt chart can function exactly like a regular chart

   
Other planned excel sheets:
1) self updating nested price list
2) self updating nested costs list
3) self updating nested Method of Manufacturing/Bill of materials List
4) welcome to other ideas!
