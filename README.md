![](https://img.shields.io/badge/STATUS-NOT%20CURRENTLY%20MAINTAINED-red.svg?longCache=true&style=flat)

# Important Notice
We have decided to stop the maintenance of this public GitHub repository.

SAP Lumira Data Access Extension for SQL Server Analysis Services (SSAS)
===========================
By [Didier Mazoue](http://scn.sap.com/people/didier.mazoue)

<strong>Lumira DA extension for SQL Server Analysis Services definition</strong> <br>
SSAS stands for SQL Server Analysis Services. The application supports Microsoft Analysis Services 2008 and 2012.<br>
SSAS is the multidimensional database provided by Microsoft.<br><br>
<strong>Additional DLLs</strong><br>
To use Lumira DA extension SQL Server Analysis Services you need to download 2 Microsoft DLLs:<br>
· Interop.ADODB.dll<br>
· Interop.ADOMD.dll<br>
<br>

<strong>Install and deploy Lumira extensions</strong><br>
1. Navigate to C:\Program Files\SAP Lumira\Desktop.<br>
2. Create the daextensions folder.<br>
3. In C:\Program Files\SAP Lumira\Desktop, open SAPLumira.ini file.<br>
4. Add the following entries to the SAPLumira.ini file, and save the file.<br>
· -Dhilo.externalds.folder=C:\Program Files\SAP Lumira\Desktop\daextensions<br>
· -Dactivate.externaldatasource.ds=true<br>
5. Copy the application “SSAS Query.exe” and the 2 DLLs into C:\Program Files\SAP Lumira\Desktop\daextensions<br><br>

<strong>Use Lumira DA extension for SQL Server Analysis Services</strong><br>
You have to connect to a SSAS server with your own credentials: click “Login” button.<br>
![My image](https://github.com/SAP/lumira-extension-da-sql-server-analysis-services-ssas/blob/master/readmescreenshots/1.png)<br>

Once connected, you select a catalog and enter an MDX query.<br>
![My image](https://github.com/SAP/lumira-extension-da-sql-server-analysis-services-ssas/blob/master/readmescreenshots/2.png)<br>
You can use Information Design Tool (delivered in SAP BI 4.x) to run a query and copy/paste the generated MDX.<br>

It’s possible to preview the results of the query in the application: click on “View query results” button.<br>
![My image](https://github.com/SAP/lumira-extension-da-sql-server-analysis-services-ssas/blob/master/readmescreenshots/3.png)<br>

Then you have to click “Run Query” button to execute the query and send the results to Lumira.<br>
![My image](https://github.com/SAP/lumira-extension-da-sql-server-analysis-services-ssas/blob/master/readmescreenshots/4.png)<br>

Files
-----------
* `Lumira SSAS Query.mp4` - Tutorial video to help users to install, deploy and use the extension.
