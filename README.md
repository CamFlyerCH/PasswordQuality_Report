## PasswordQuality Report

These two scripts help to create a nice report from the output of [DSInternals](https://github.com/MichaelGrafnetter/DSInternals/blob/master/Documentation/PowerShell/Readme.md) [Test-PasswordQuality](https://www.dsinternals.com/en/auditing-active-directory-password-quality/) function.

There are 2 scripts, because I think the procedure should be done in 2 steps:

1.  Data-Collection on the Domain Controller with the installed DSInternals PowerShell Module
2.  Creation of the Excel report from the collected data with the help of the excelent [ImportExcel](https://github.com/dfinke/ImportExcel) PowerShell module to import/export Excel spreadsheets, without the need of Excel itself.

Check the Demo folder for a sample of the created files.
<br/><br/>

### Step by step guide

#### 1 - Download an updated list of leakad passwords (hashes)

To get this list (about a 30 GByte text file) from [have i been pwned](https://haveibeenpwned.com/Passwords) use the [PawnedPasswordsDownloader](https://github.com/HaveIBeenPwned/PwnedPasswordsDownloader).
<br/><br/>

#### 2 - Install the DSInternals PowerShell module on one Domain Controller

Install the module [DSInternals](https://github.com/MichaelGrafnetter/DSInternals/blob/master/Documentation/PowerShell/Readme.md) directly with `Install-Module DSInternals -Force`  or by saving it on a workstation first with for example `Save-Module DSInternals -Path C:\Data\PSModules` and then copy it to the domain controller to the path `C:\Windows\System32\WindowsPowerShell\v1.0\Modules` .
<br/><br/>

#### 3 - Put the script and the hash file on the domain controller

Create a folder (for Example D:\\PW-Audit) on the Domain Controller on a disk with sufficient space for the has file.   
Copy the hash file download in step 1 to this folder or an a network share accessible by the Domain Controller (with high performance).  
Also download the script [PasswordQuality-Create-Exports.ps1](https://github.com/CamFlyerCH/PasswordQuality_Report/raw/main/PasswordQuality-Create-Exports.ps1) from this repository and copy it to the above created folder.
<br/><br/>

#### 4 - Modify and execute the Export-Script

Open the Script PasswordQuality-Create-Exports.ps1 in an editor an change line 4 (`$SortedHashFile = "D:\PW-Audit\pwnedpasswords_ntlm.txt"`) to represent the actual path to the hashes file. Then execute the script. This will take a few minutes while searching hashes in the big file.
<br/><br/>

#### 5 - Install the ImportExcel Module on your computer of choice

To create the Excel sheet, you need to have the PowerShell Module [ImportExcel](https://github.com/dfinke/ImportExcel) installed on the workstation where you also what to view the report later. The module is also available from the [Powershell Gallery](https://www.powershellgallery.com/packages/ImportExcel). The simplest way to install the module will be to execute `Install-Module -Name ImportExcel -Force` in an (elevated) PowerShell window.
<br/><br/>

#### 6 - Download the Create-Report script to a folder

Download the second script [PasswordQuality-Create-Report.ps1](https://github.com/CamFlyerCH/PasswordQuality_Report/raw/main/PasswordQuality-Create-Report.ps1) from this repository to a folder where you plan to have the report. Depending on you PowerShell configuration, this folder should be localy. Also don't forget to unblock it after downloading it from the internet.
<br/><br/>

#### 7 - Transfer the exports to the folder on the workstation

The export script run in point 4 should create three files:  
PasswordQuality\_\<DomainName>\_Accounts.xml  
PasswordQuality\_\<DomainName>\_PWQ-Data.txt  
PasswordQuality\_\<DomainName>\_PWQ-Data.xml

Copy at least the two .xml files to the same folder as the PasswordQuality-Create-Report.ps1 script.
<br/><br/>

#### 8 - Execute the Create-Report script

The script will first look for all PWQ-Data.xml files and then work on these files for every domain name found.  
It will create and open an Excel file named `PasswordQuality_<DomainName>_Report_<ExportModifyDate>.xlsx` 

If Excel is not installed on the machine, an requester will pop up, but the file is created anyways.
