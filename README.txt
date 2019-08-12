This workbook serves only for demonstration purposes of my work only.

Note:
This is one of the tools that I created at my workplace to keep things organized when we need to create new orgcode (ID value) for a new company. It is used to avoid duplicities in our database.
I have removed sensitive information from the code. Due to this fact, macros won't work if you try to run them.

Description:
- On Sheet("Sheet1") you can see 2 tables. The table on the left normally contains all existing orgcodes and companies that we already have in the DB. This would normally be updated once per year. In the right table, we would write new codes that we are creating. Main information is client_org_code (ID) and legal_name (company name), the rest is just to identify more internal things.

- Sheet("Checks") would normally be hidden and serves only for macro functionality + internal checks if needed. There are 3 tables filled from DB using SQL by a macro.
	 - Old codes - list of org codes from Sheet1 left the table - it serves for comparison if all codes in Sheet1 left the table are existing in the DB and if the legal_name matches the DB
	 - New codes -  list of org codes from Sheet1 right table - it serves for comparison if all codes in Sheet1 right table were created in the DB and if the legal_name matches the DB
	 - Possible new codes - when we create new codes we always add 1 e.g. 1234,1235,1236 etc. This is list based on codes higher than the last code indicated in Sheet1 right table - it's purpose 	is to discover when new code was created, but not indicated in this workbook.

- On Sheet1, in the right top corner, there are 2 buttons
 	- Check downloads DB information through SQL and compares it against Sheet1 tables as described above.
	- Back-up simply creates copy and saves it in specified locations.
	- Below those buttons, there are dates updated when macros are run to keep track of last check and back-up

- Checks are indicated in columns B and M. Those are hidden untill some issue appears - e.g. legal_name missmatch.

- I have included some simple whitelist of Windows users in macros as well, so only certain people in the company can use those macros, others won't be able to. Normally workbook + worksheet are locked + VBA is locked, therefore standard user wouldn't be able to get over this. I am aware that Worksheet/Workbook/VBA protections can be easily removed and bypassed, but for internal uses, in this case, it is not an issue.
