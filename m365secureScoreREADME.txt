resource - https://gcits.com/knowledge-base/export-customers-microsoft-secure-scores-to-csv-and-html-reports/

All credit to author - Elliot Munro


1. Add a link to your logo into the $yourLogo variable.

2. Modify the $homePage and $logoutURI values to any valid URI that you like. They don’t need to be actual addresses, so feel free to make something up. Set the $appIDUri variable to a use a valid domain in your tenant. eg. https://yourdomain.com/$((New-Guid).ToString())

3. If you are running the script for a single tenant, comment out the following two lines:

4. Press F5 to run the script
5. Sign in to Azure AD using your global admin credentials. Note that the login window may appear behind Visual Studio Code.
6. Wait for the script to complete.

7. You can find the exported HTML reports and CSV overview at C:\temp\SecureScoreReports\
